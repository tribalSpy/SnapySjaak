import React, { useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import "./styles.css";

const REFRESH_INTERVAL_MS = 15000;
const CMR_A4_WIDTH = 827;
const CMR_A4_HEIGHT = 1169;
const CMR_DOCUMENT_PADDING = 20;
const CMR_DEFAULT_FIELD_WIDTH = 140;
const CMR_DEFAULT_FIELD_HEIGHT = 52;
const PERMISSIONS = {
  PHOTOS_VIEW: "photos:view",
  FUST_VIEW: "fust:view",
  FUST_IN: "fust:in",
  FUST_OUT: "fust:out",
  FUST_OVERVIEW: "fust:overview",
  FUST_MANAGE: "fust:manage",
  CMR_VIEW: "cmr:view",
  HAL_LOCATIONS_VIEW: "hal_locations:view",
  EXPEDITION_STICKERS_VIEW: "expedition_stickers:view",
  BUNCHES_VIEW: "bunches:view",
  DAG_FOUTJES_VIEW: "expedition_stickers:view",
  FOUTEN_OVERVIEW_VIEW: "fouten_overview:view",
  CMR_MANAGE: "cmr:manage",
  CLOCK_VIEW: "clock:view",
  CLOCK_MANAGE: "clock:manage",
  USERS_MANAGE: "users:manage",
  SETTINGS_MANAGE: "settings:manage",
  UKDOCS_VIEW: "ukdocs:view",
  UKDOCS_INSPECTION_VIEW: "ukdocs_inspection:view",
};
const ALL_PERMISSIONS = Object.values(PERMISSIONS);
const DEFAULT_PERMISSIONS_BY_ROLE = {
  admin: ALL_PERMISSIONS,
  viewer: [PERMISSIONS.PHOTOS_VIEW],
};
const PAGE_DEFINITIONS = [
  { key: "dashboard", label: "Photos", permission: PERMISSIONS.PHOTOS_VIEW },
  { key: "fust", label: "Fust", permission: PERMISSIONS.FUST_VIEW },
  { key: "cmrprint", label: "CMR Print", permission: PERMISSIONS.CMR_VIEW },
  { key: "hallocations", label: "Hal Locations", permission: PERMISSIONS.HAL_LOCATIONS_VIEW },
  { key: "expeditionstickers", label: "Expedition Sticker", permission: PERMISSIONS.EXPEDITION_STICKERS_VIEW },
  { key: "bunches", label: "Bunches", permission: PERMISSIONS.BUNCHES_VIEW },
  { key: "dagfoutjes", label: "Fout Registratie", permission: PERMISSIONS.DAG_FOUTJES_VIEW },
  { key: "foutenoverzicht", label: "Fouten Overzicht", permission: PERMISSIONS.FOUTEN_OVERVIEW_VIEW },
  { key: "ukdocsprint", label: "UKdocs Print", permission: PERMISSIONS.UKDOCS_VIEW },
  { key: "ukdocsinspection", label: "Phyto Inspection", permission: PERMISSIONS.UKDOCS_INSPECTION_VIEW },
  { key: "clock", label: "Inklokken", permission: PERMISSIONS.CLOCK_VIEW },
  { key: "users", label: "Users", permission: PERMISSIONS.USERS_MANAGE },
  { key: "settings", label: "Settings", permission: PERMISSIONS.SETTINGS_MANAGE },
  { key: "ukdocs", label: "UKdocs", permission: PERMISSIONS.UKDOCS_VIEW },
];

function formatTimestamp(value) {
  if (!value) {
    return "unknown time";
  }

  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) {
    return value;
  }

  return parsed.toLocaleString();
}

function localDateIso() {
  const now = new Date();
  const offsetMinutes = now.getTimezoneOffset();
  return new Date(now.getTime() - (offsetMinutes * 60 * 1000)).toISOString().slice(0, 10);
}

function normalizePermissions(role, permissions) {
  if (role === "admin") {
    return [...DEFAULT_PERMISSIONS_BY_ROLE.admin];
  }
  const defaults = DEFAULT_PERMISSIONS_BY_ROLE[role] || DEFAULT_PERMISSIONS_BY_ROLE.viewer;
  if (!Array.isArray(permissions)) {
    return [...defaults];
  }

  const normalized = [...new Set(
    permissions
      .map((value) => String(value || "").trim())
      .filter((value) => ALL_PERMISSIONS.includes(value)),
  )];

  return normalized.length ? normalized : [...defaults];
}

function hasPermission(user, permission) {
  if (!user) {
    return false;
  }

  const permissions = normalizePermissions(user.role, user.permissions);
  return permissions.includes(permission);
}

function availablePagesForUser(user) {
  return PAGE_DEFINITIONS.filter((page) => hasPermission(user, page.permission));
}

function defaultPageForUser(user) {
  return availablePagesForUser(user)[0]?.key || "dashboard";
}

function pageHeading(page) {
  switch (page) {
    case "fust":
      return {
        title: "Fust Management",
        caption: "Capture IN and OUT movements, import outside actions, and review balances with control and beheer.",
      };
    case "clock":
      return {
        title: "Inklokken",
        caption: "Scan badge codes, add manual corrections, and export clocked times online.",
      };
    case "hallocations":
      return {
        title: "Hal Locations",
        caption: "Upload a halindeling and generate the same sticker PDF flow directly inside the Shadow app.",
      };
    case "expeditionstickers":
      return {
        title: "Expedition Sticker",
        caption: "",
      };
    case "bunches":
      return {
        title: "Bunches",
        caption: "Process bunches exports, keep article and APE master data, and reuse saved runs across the team.",
      };
    case "dagfoutjes":
      return {
        title: "Fout Registratie",
        caption: "",
      };
    case "foutenoverzicht":
      return {
        title: "Fouten Overzicht",
        caption: "Review mistakes by person, type, day, week, and month in a separate protected overview.",
      };
    case "cmrprint":
      return {
        title: "CMR Print",
        caption: "Use your imported CMR customers, linked profiles, and saved templates in a separate daily workspace.",
      };
    case "users":
      return {
        title: "Users",
        caption: "Control which menus and actions each account can use.",
      };
    case "settings":
      return {
        title: "Settings",
        caption: "Prepare email recipients, spreadsheet mapping, and business master data.",
      };
    case "ukdocs":
      return {
        title: "UKdocs",
        caption: "Prepare UK export shipments, saved mappings, audit checks, and document generation workflows inside Shadow App.",
      };
    case "ukdocsprint":
      return {
        title: "UKdocs Print",
        caption: "Collect the extra export files per finished UK shipment and keep everything together by invoice and truck registration.",
      };
    case "ukdocsinspection":
      return {
        title: "Phyto Inspection",
        caption: "",
      };
    default:
      return {
        title: "Sjaak vd Vijver Expedition Photo Dashboard",
        caption: "Choose departure date and optionally filter by customer code.",
      };
  }
}

function imageUrl(image, run, retryKey = "") {
  const params = new URLSearchParams();
  params.set("id", image.id);
  params.set("account", String(run?.metadata?.drive_account || "default"));
  if (image.mime_type) {
    params.set("mime", image.mime_type);
  }
  if (retryKey) {
    params.set("retry", String(retryKey));
  }
  return `/api/image?${params.toString()}`;
}

function PhotoImage({ image, run, alt, loading = "eager" }) {
  const [attempt, setAttempt] = useState(0);
  const [failed, setFailed] = useState(false);

  useEffect(() => {
    setAttempt(0);
    setFailed(false);
  }, [image.id, run?.folder_id]);

  useEffect(() => {
    if (!failed || attempt >= 3) {
      return undefined;
    }

    const timer = window.setTimeout(() => {
      setFailed(false);
      setAttempt((value) => value + 1);
    }, 1200);

    return () => window.clearTimeout(timer);
  }, [attempt, failed]);

  return (
    <img
      src={imageUrl(image, run, attempt)}
      alt={alt}
      loading={loading}
      onError={() => setFailed(true)}
    />
  );
}

async function apiJson(path, options = {}) {
  const response = await fetch(path, {
    ...options,
    headers: {
      ...(options.body ? { "content-type": "application/json" } : {}),
      ...(options.headers || {}),
    },
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(payload.error || `Request failed with ${response.status}`);
  }
  return payload;
}

function emptyFustMetrics() {
  return { dc: 0, cctag: 0, dcs: 0, dco: 0, pal: 0, vk: 0 };
}

function DocumentStatus({ action }) {
  const [showError, setShowError] = useState(false);
  const documentKind = action?.type === "IN" ? "fustbon" : "cmr";
  const document = action?.type === "IN" ? action?.fustbon || {} : action?.cmr || {};
  const label = action?.type === "IN" ? "Fustbon" : "CMR";
  if (!action || !["IN", "OUT"].includes(action.type)) {
    return "-";
  }
  if (document.status === "uploaded" && document.file_id) {
    const href = `/api/fust/actions/${encodeURIComponent(action.id)}/document/${documentKind}`;
    return <a href={href} target="_blank" rel="noreferrer">Open {label}</a>;
  }
  if (document.status === "uploaded" && document.web_link) {
    return <a href={document.web_link} target="_blank" rel="noreferrer">Open {label}</a>;
  }
  if (document.status === "skipped") {
    return "Skipped";
  }
  if (document.status === "failed") {
    return (
      <div>
        <button type="button" className="link-button" onClick={() => setShowError((value) => !value)}>
          Not available
        </button>
        {showError && <div>{document.error || "Upload failed"}</div>}
      </div>
    );
  }
  return "Not available";
}


function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || "").split(",").pop() || "");
    reader.onerror = () => reject(reader.error || new Error("Could not read file"));
    reader.readAsDataURL(file);
  });
}

function downloadBase64File(filename, contentBase64, mimeType) {
  const binary = window.atob(contentBase64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  const blob = new Blob([bytes], { type: mimeType || "application/octet-stream" });
  const url = window.URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = safeDownloadFilename(filename);
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.URL.revokeObjectURL(url);
}

function safeDownloadFilename(filename) {
  const cleaned = String(filename || "download")
    .replace(/[\\/:*?"<>|]/g, "-")
    .replace(/\s+/g, " ")
    .trim();
  return cleaned || "download";
}

function HalLocationsPage() {
  const [file, setFile] = useState(null);
  const [sessionId, setSessionId] = useState("");
  const [locPrefixes, setLocPrefixes] = useState([]);
  const [custPrefixes, setCustPrefixes] = useState([]);
  const [visibleCustPrefixes, setVisibleCustPrefixes] = useState([]);
  const [custByLoc, setCustByLoc] = useState({});
  const [selectedLocPrefixes, setSelectedLocPrefixes] = useState([]);
  const [selectedCustPrefixes, setSelectedCustPrefixes] = useState([]);
  const [uploadMessage, setUploadMessage] = useState("");
  const [uploadError, setUploadError] = useState("");
  const [generateMessage, setGenerateMessage] = useState("");
  const [generateError, setGenerateError] = useState("");
  const [uploading, setUploading] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [loadingSheet, setLoadingSheet] = useState(false);
  const [sourceLabel, setSourceLabel] = useState("");

  function toggleValue(value, selectedValues, setSelectedValues) {
    setSelectedValues(
      selectedValues.includes(value)
        ? selectedValues.filter((item) => item !== value)
        : [...selectedValues, value],
    );
  }

  function replaceSelection(values, setSelectedValues) {
    setSelectedValues([...values]);
  }

  async function handleUpload() {
    if (!file) {
      setUploadError("Select a file first");
      setUploadMessage("");
      return;
    }

    setUploading(true);
    setUploadError("");
    setUploadMessage("Uploading halindeling...");
    setGenerateError("");
    setGenerateMessage("");

    try {
      const contentBase64 = await fileToBase64(file);
      const payload = await apiJson("/api/hal-locations/inspect", {
        method: "POST",
        body: JSON.stringify({
          file: {
            name: file.name,
            content_base64: contentBase64,
          },
        }),
      });

      setSessionId(payload.id || "");
      setLocPrefixes(Array.isArray(payload.locPrefixes) ? payload.locPrefixes : []);
      setCustPrefixes(Array.isArray(payload.custPrefixes) ? payload.custPrefixes : []);
      setVisibleCustPrefixes(Array.isArray(payload.custPrefixes) ? payload.custPrefixes : []);
      setCustByLoc(payload.custByLoc || {});
      setSelectedLocPrefixes([]);
      setSelectedCustPrefixes([]);
      setSourceLabel(payload.source?.type === "upload" ? `Source: uploaded file ${payload.source?.file_name || file.name}` : "");
      setUploadMessage(`Ready: ${payload.totalRows || 0} rows, ${(payload.locPrefixes || []).length} location prefixes, ${(payload.custPrefixes || []).length} customer prefixes.`);
    } catch (error) {
      setSessionId("");
      setLocPrefixes([]);
      setCustPrefixes([]);
      setVisibleCustPrefixes([]);
      setCustByLoc({});
      setSelectedLocPrefixes([]);
      setSelectedCustPrefixes([]);
      setUploadError(error.message);
      setUploadMessage("");
    } finally {
      setUploading(false);
    }
  }

  async function handleLoadFromSheet() {
    setLoadingSheet(true);
    setUploadError("");
    setUploadMessage("Loading ERP_PASTE from Google Sheets...");
    setGenerateError("");
    setGenerateMessage("");

    try {
      const payload = await apiJson("/api/hal-locations/load-sheet", {
        method: "POST",
        body: JSON.stringify({}),
      });

      setSessionId(payload.id || "");
      setLocPrefixes(Array.isArray(payload.locPrefixes) ? payload.locPrefixes : []);
      setCustPrefixes(Array.isArray(payload.custPrefixes) ? payload.custPrefixes : []);
      setVisibleCustPrefixes(Array.isArray(payload.custPrefixes) ? payload.custPrefixes : []);
      setCustByLoc(payload.custByLoc || {});
      setSelectedLocPrefixes([]);
      setSelectedCustPrefixes([]);
      setSourceLabel(payload.source?.sheet_name ? `Source: ${payload.source.sheet_name} (${payload.source.spreadsheet_id || "spreadsheet"})` : "Source: ERP_PASTE");
      setUploadMessage(`Ready: ${payload.totalRows || 0} rows, ${(payload.locPrefixes || []).length} location prefixes, ${(payload.custPrefixes || []).length} customer prefixes.`);
    } catch (error) {
      setSessionId("");
      setLocPrefixes([]);
      setCustPrefixes([]);
      setVisibleCustPrefixes([]);
      setCustByLoc({});
      setSelectedLocPrefixes([]);
      setSelectedCustPrefixes([]);
      setSourceLabel("");
      setUploadError(error.message);
      setUploadMessage("");
    } finally {
      setLoadingSheet(false);
    }
  }

  function filterCustomersBySelectedLocations() {
    if (!selectedLocPrefixes.length) {
      setGenerateError("Select at least one location prefix first");
      setGenerateMessage("");
      return;
    }

    const allowed = [...new Set(
      selectedLocPrefixes.flatMap((prefix) => Array.isArray(custByLoc[prefix]) ? custByLoc[prefix] : []),
    )].sort((left, right) => left.localeCompare(right));

    setVisibleCustPrefixes(allowed);
    setSelectedCustPrefixes([]);
    setGenerateError("");
    setGenerateMessage(`${allowed.length} customer prefixes matched the selected locations.`);
  }

  async function handleGenerate() {
    if (!sessionId) {
      setGenerateError("Upload a halindeling first");
      setGenerateMessage("");
      return;
    }
    if (!selectedLocPrefixes.length) {
      setGenerateError("Select at least one location prefix");
      setGenerateMessage("");
      return;
    }

    setGenerating(true);
    setGenerateError("");
    setGenerateMessage("Generating sticker PDF...");

    try {
      const response = await fetch("/api/hal-locations/generate", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({
          id: sessionId,
          locPrefixes: selectedLocPrefixes,
          custPrefixes: selectedCustPrefixes,
        }),
      });

      if (!response.ok) {
        const payload = await response.json().catch(() => ({}));
        throw new Error(payload.error || `Request failed with ${response.status}`);
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `stickers_${new Date().toISOString().slice(0, 10)}.pdf`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);
      setGenerateMessage("Sticker PDF downloaded.");
    } catch (error) {
      setGenerateError(error.message);
      setGenerateMessage("");
    } finally {
      setGenerating(false);
    }
  }

  return (
    <section className="hal-page">
      <article className="panel hal-panel">
        <div className="section-header">
          <div>
            <h2>Upload halindeling</h2>
            <p>Load the data directly from the configured Google Sheet tab `ERP_PASTE`, or fall back to the old Excel upload when needed.</p>
          </div>
        </div>
        <div className="hal-upload-row">
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={(event) => setFile(event.target.files?.[0] || null)}
          />
          <button type="button" onClick={handleLoadFromSheet} disabled={loadingSheet}>
            {loadingSheet ? "Loading sheet..." : "Load from Hal Indeling Spreadsheet"}
          </button>
          <button type="button" className="primary" onClick={handleUpload} disabled={uploading}>
            {uploading ? "Uploading..." : "Upload file"}
          </button>
        </div>
        {sourceLabel ? <div className="notice">{sourceLabel}</div> : null}
        {uploadMessage ? <div className="notice success">{uploadMessage}</div> : null}
        {uploadError ? <div className="notice danger">{uploadError}</div> : null}
      </article>

      <article className="panel hal-panel">
        <div className="section-header">
          <div>
            <h2>Select locations</h2>
            <p>These are the first 2 characters of each location code, like <code>gK</code>, <code>gL</code>, <code>bA</code>, and <code>eT</code>.</p>
          </div>
        </div>
        <div className="row-actions hal-actions-row">
          <button type="button" onClick={() => replaceSelection(locPrefixes, setSelectedLocPrefixes)} disabled={!locPrefixes.length}>Select all</button>
          <button type="button" onClick={() => replaceSelection([], setSelectedLocPrefixes)} disabled={!locPrefixes.length}>Clear</button>
        </div>
        <div className="hal-chip-grid">
          {locPrefixes.map((prefix) => (
            <label key={prefix} className={`hal-chip ${selectedLocPrefixes.includes(prefix) ? "selected" : ""}`}>
              <input
                type="checkbox"
                checked={selectedLocPrefixes.includes(prefix)}
                onChange={() => toggleValue(prefix, selectedLocPrefixes, setSelectedLocPrefixes)}
              />
              <span>{prefix}</span>
            </label>
          ))}
          {!locPrefixes.length ? <div className="notice">Upload a halindeling to load location prefixes.</div> : null}
        </div>
      </article>

      <article className="panel hal-panel">
        <div className="section-header">
          <div>
            <h2>Select customer prefixes</h2>
            <p>Number-first customer codes use 3 characters, letter-first customer codes use 2. Leave empty to include every customer on the chosen locations.</p>
          </div>
        </div>
        <div className="row-actions hal-actions-row">
          <button type="button" onClick={() => replaceSelection(visibleCustPrefixes, setSelectedCustPrefixes)} disabled={!visibleCustPrefixes.length}>Select all</button>
          <button type="button" onClick={() => replaceSelection([], setSelectedCustPrefixes)} disabled={!visibleCustPrefixes.length}>Clear</button>
          <button type="button" onClick={filterCustomersBySelectedLocations} disabled={!locPrefixes.length}>Only prefixes on selected locations</button>
          <button type="button" onClick={() => { setVisibleCustPrefixes(custPrefixes); setSelectedCustPrefixes([]); setGenerateError(""); setGenerateMessage("All customer prefixes restored."); }} disabled={!custPrefixes.length}>Show all prefixes</button>
        </div>
        <div className="hal-chip-grid">
          {visibleCustPrefixes.map((prefix) => (
            <label key={prefix} className={`hal-chip ${selectedCustPrefixes.includes(prefix) ? "selected" : ""}`}>
              <input
                type="checkbox"
                checked={selectedCustPrefixes.includes(prefix)}
                onChange={() => toggleValue(prefix, selectedCustPrefixes, setSelectedCustPrefixes)}
              />
              <span>{prefix}</span>
            </label>
          ))}
          {!visibleCustPrefixes.length ? <div className="notice">No customer prefixes available for the current selection.</div> : null}
        </div>
      </article>

      <article className="panel hal-panel">
        <div className="section-header">
          <div>
            <h2>Generate PDF</h2>
            <p>The output matches the StickerPrinter layout: one sticker per unique customer, 10 x 15 cm, rotated, with the location large and the customer code below.</p>
          </div>
        </div>
        <div className="row-actions hal-actions-row">
          <button type="button" className="primary" onClick={handleGenerate} disabled={generating || !sessionId || !selectedLocPrefixes.length}>
            {generating ? "Generating..." : "Download stickers PDF"}
          </button>
        </div>
        {generateMessage ? <div className="notice success">{generateMessage}</div> : null}
        {generateError ? <div className="notice danger">{generateError}</div> : null}
      </article>
    </section>
  );
}

function ExpeditionStickerPage() {
  const [planningFile, setPlanningFile] = useState(null);
  const [splitFile, setSplitFile] = useState(null);
  const [savedState, setSavedState] = useState(null);
  const [loadingState, setLoadingState] = useState(true);
  const [savingSources, setSavingSources] = useState(false);
  const [sheetBusy, setSheetBusy] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [halSessionId, setHalSessionId] = useState("");
  const [halSummary, setHalSummary] = useState(null);
  const [generatedFiles, setGeneratedFiles] = useState([]);
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");

  async function loadState() {
    setLoadingState(true);
    try {
      const payload = await apiJson("/api/expedition-stickers");
      setSavedState(payload);
      setError("");
    } catch (loadError) {
      setError(loadError.message);
    } finally {
      setLoadingState(false);
    }
  }

  useEffect(() => {
    loadState();
  }, []);

  const hasSavedSources = Boolean(savedState?.planning_file || savedState?.split_file);

  async function clearSavedSource(kind) {
    setSavingSources(true);
    setError("");
    setMessage("");
    try {
      const payload = await apiJson(`/api/expedition-stickers/source/${kind}`, {
        method: "DELETE",
      });
      setSavedState((current) => ({
        ...(current || {}),
        ...payload,
      }));
      if (kind === "planning") {
        setPlanningFile(null);
      } else if (kind === "split") {
        setSplitFile(null);
      }
      setGeneratedFiles([]);
      setHalSessionId("");
      setHalSummary(null);
      setMessage(`${kind === "planning" ? "Planning" : "Split"} file removed.`);
    } catch (deleteError) {
      setError(deleteError.message);
      setMessage("");
    } finally {
      setSavingSources(false);
    }
  }

  async function saveSources() {
    if (!planningFile && !splitFile) {
      setError("Choose a planning file or split file first");
      setMessage("");
      return;
    }

    setSavingSources(true);
    setError("");
    setMessage("");
    try {
      const body = {};
      if (planningFile) {
        body.planning_file = {
          name: planningFile.name,
          content_base64: await fileToBase64(planningFile),
        };
      }
      if (splitFile) {
        body.split_file = {
          name: splitFile.name,
          content_base64: await fileToBase64(splitFile),
        };
      }
      const payload = await apiJson("/api/expedition-stickers/upload", {
        method: "POST",
        body: JSON.stringify(body),
      });
      setSavedState((current) => ({
        ...(current || {}),
        ...payload,
      }));
      setPlanningFile(null);
      setSplitFile(null);
      setGeneratedFiles([]);
      setHalSessionId("");
      setHalSummary(null);
      setMessage("Source files saved.");
    } catch (saveError) {
      setError(saveError.message);
      setMessage("");
    } finally {
      setSavingSources(false);
    }
  }

  async function loadHalindelingFromSheet() {
    setSheetBusy(true);
    setError("");
    setMessage("Loading ERP_PASTE from Google Sheets...");
    setGeneratedFiles([]);
    try {
      const payload = await apiJson("/api/expedition-stickers/load-sheet", {
        method: "POST",
        body: JSON.stringify({}),
      });
      setHalSessionId(payload.id || "");
      setHalSummary(payload);
      setMessage(`ERP_PASTE loaded: ${payload.totalRows || 0} rows, ${(payload.locPrefixes || []).length} location prefixes.`);
      return payload;
    } catch (sheetError) {
      setHalSessionId("");
      setHalSummary(null);
      setError(sheetError.message);
      setMessage("");
      return null;
    } finally {
      setSheetBusy(false);
    }
  }

  async function generateExpeditionStickers(nextSessionId = "") {
    const activeSessionId = nextSessionId || halSessionId;
    if (!activeSessionId) {
      setError("Load ERP_PASTE first");
      setMessage("");
      return;
    }

    setGenerating(true);
    setError("");
    setMessage("Generating expedition sticker PDFs...");
    try {
      const payload = await apiJson("/api/expedition-stickers/generate", {
        method: "POST",
        body: JSON.stringify({ id: activeSessionId }),
      });
      setGeneratedFiles(payload.files || []);
      const missing = Array.isArray(payload.summary?.missing_locations) ? payload.summary.missing_locations : [];
      setMessage(
        `Generated ${(payload.files || []).length} PDF file(s) from ${payload.summary?.combined_row_count || 0} sticker rows.${missing.length ? ` Missing hal locations for: ${missing.join(", ")}` : ""}`,
      );
    } catch (generateError) {
      setError(generateError.message);
      setMessage("");
    } finally {
      setGenerating(false);
    }
  }

  async function handleSourceFileChange(nextPlanningFile, nextSplitFile) {
    setPlanningFile(nextPlanningFile);
    setSplitFile(nextSplitFile);
    if (!nextPlanningFile && !nextSplitFile) {
      return;
    }
    setSavingSources(true);
    setError("");
    setMessage("");
    try {
      const body = {};
      if (nextPlanningFile) {
        body.planning_file = {
          name: nextPlanningFile.name,
          content_base64: await fileToBase64(nextPlanningFile),
        };
      }
      if (nextSplitFile) {
        body.split_file = {
          name: nextSplitFile.name,
          content_base64: await fileToBase64(nextSplitFile),
        };
      }
      const payload = await apiJson("/api/expedition-stickers/upload", {
        method: "POST",
        body: JSON.stringify(body),
      });
      setSavedState((current) => ({
        ...(current || {}),
        ...payload,
      }));
      setPlanningFile(null);
      setSplitFile(null);
      setGeneratedFiles([]);
      setHalSessionId("");
      setHalSummary(null);
      setMessage("Source files saved.");
    } catch (saveError) {
      setError(saveError.message);
      setMessage("");
    } finally {
      setSavingSources(false);
    }
  }

  async function handleContinue() {
    if (!hasSavedSources && !planningFile && !splitFile) {
      setError("Upload at least one source file first");
      setMessage("");
      return;
    }
    const sheetPayload = await loadHalindelingFromSheet();
    if (!sheetPayload?.id) {
      return;
    }
    await generateExpeditionStickers(sheetPayload.id);
  }

  return (
    <section className="overview-stack">
      <article className="panel hal-panel">
        <div className="section-header">
          <div>
            <h2>Source files</h2>
          </div>
          <button type="button" onClick={loadState} disabled={loadingState}>
            {loadingState ? "Refreshing..." : "Refresh saved state"}
          </button>
        </div>
        <div className="form-grid">
          <label>
            <span>Planning file</span>
            <input type="file" accept=".xlsx,.xls,.csv" onChange={(event) => handleSourceFileChange(event.target.files?.[0] || null, null)} />
          </label>
          <label>
            <span>Split file</span>
            <input type="file" accept=".xlsx,.xls,.csv" onChange={(event) => handleSourceFileChange(null, event.target.files?.[0] || null)} />
          </label>
        </div>
        {savingSources ? <div className="notice">Saving source files...</div> : null}
        <div className="table-wrap">
          <table className="data-table">
            <thead>
              <tr>
                <th>Source</th>
                <th>Saved file</th>
                <th>Rows</th>
                <th>Saved by</th>
                <th>Saved at</th>
                <th>Action</th>
              </tr>
            </thead>
            <tbody>
              {[
                ["Planning", "planning", savedState?.planning_file, savedState?.planning_summary],
                ["Split", "split", savedState?.split_file, savedState?.split_summary],
              ].map(([label, kind, file, summary]) => (
                <tr key={label}>
                  <td>{label}</td>
                  <td>{file?.original_name || "-"}</td>
                  <td>{summary?.row_count || 0}</td>
                  <td>{file?.saved_by || "-"}</td>
                  <td>{file?.saved_at ? formatTimestamp(file.saved_at) : "-"}</td>
                  <td>
                    <button type="button" onClick={() => clearSavedSource(kind)} disabled={!file || savingSources}>
                      Remove
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </article>

      <article className="panel hal-panel">
        <div className="section-header">
          <div>
            <h2>Continue</h2>
          </div>
        </div>
        <div className="row-actions hal-actions-row">
          <button type="button" className="primary" onClick={handleContinue} disabled={savingSources || sheetBusy || generating || !hasSavedSources}>
            {sheetBusy ? "Loading ERP_PASTE..." : generating ? "Generating..." : "Continue"}
          </button>
          <button type="button" disabled={!generatedFiles.length} onClick={() => generatedFiles.forEach((file) => downloadBase64File(file.name, file.content_base64, file.mime_type))}>
            Download all
          </button>
        </div>
        {halSummary ? <p className="sidebar-note">ERP_PASTE rows: {halSummary.totalRows || 0}</p> : null}
        {!!generatedFiles.length && (
          <div className="row-actions hal-actions-row">
            {generatedFiles.map((file) => (
              <button key={file.name} type="button" onClick={() => downloadBase64File(file.name, file.content_base64, file.mime_type)}>
                {file.name}
              </button>
            ))}
          </div>
        )}
      </article>

      {message ? <div className="notice success">{message}</div> : null}
      {error ? <div className="notice danger">{error}</div> : null}
    </section>
  );
}

function useFustMeta(enabled) {
  const [state, setState] = useState({ loading: true, data: null, error: "" });

  useEffect(() => {
    if (!enabled) {
      setState({ loading: false, data: null, error: "" });
      return undefined;
    }

    let cancelled = false;
    setState((current) => ({ ...current, loading: true, error: "" }));
    apiJson("/api/fust/meta")
      .then((payload) => {
        if (!cancelled) {
          setState({ loading: false, data: payload, error: "" });
        }
      })
      .catch((error) => {
        if (!cancelled) {
          setState({ loading: false, data: null, error: error.message });
        }
      });

    return () => {
      cancelled = true;
    };
  }, [enabled]);

  return state;
}

function useFustActions(enabled) {
  const [refreshKey, setRefreshKey] = useState(0);
  const [state, setState] = useState({ loading: true, data: null, error: "" });

  useEffect(() => {
    if (!enabled) {
      setState({ loading: false, data: null, error: "" });
      return undefined;
    }

    let cancelled = false;
    setState((current) => ({ ...current, loading: true, error: "" }));
    apiJson("/api/fust/actions")
      .then((payload) => {
        if (!cancelled) {
          setState({ loading: false, data: payload, error: "" });
        }
      })
      .catch((error) => {
        if (!cancelled) {
          setState({ loading: false, data: null, error: error.message });
        }
      });

    return () => {
      cancelled = true;
    };
  }, [enabled, refreshKey]);

  return {
    ...state,
    refresh: () => setRefreshKey((value) => value + 1),
  };
}

function useDashboardData(selectedDate, searchTerm, syncVersion, enabled) {
  const [state, setState] = useState({
    data: null,
    loading: true,
    error: "",
  });

  useEffect(() => {
    if (!enabled) {
      setState({ data: null, loading: false, error: "" });
      return undefined;
    }

    const controller = new AbortController();
    const params = new URLSearchParams();
    if (selectedDate) {
      params.set("date", selectedDate);
    }
    if (searchTerm.trim()) {
      params.set("search", searchTerm.trim());
    }

    setState((current) => ({ ...current, loading: true, error: "" }));
    fetch(`/api/data?${params.toString()}`, { signal: controller.signal })
      .then((response) => {
        if (!response.ok) {
          throw new Error(`Data request failed with ${response.status}`);
        }
        return response.json();
      })
      .then((payload) => {
        setState({ data: payload, loading: false, error: "" });
      })
      .catch((error) => {
        if (error.name !== "AbortError") {
          setState({ data: null, loading: false, error: error.message });
        }
      });

    return () => controller.abort();
  }, [selectedDate, searchTerm, syncVersion, enabled]);

  return state;
}

function useSyncStatus(enabled) {
  const [status, setStatus] = useState({});

  useEffect(() => {
    if (!enabled) {
      setStatus({});
      return undefined;
    }

    let stopped = false;

    async function loadStatus() {
      try {
        const response = await fetch("/api/status");
        if (!response.ok) {
          return;
        }
        const payload = await response.json();
        if (!stopped) {
          setStatus(payload);
        }
      } catch {
        if (!stopped) {
          setStatus({});
        }
      }
    }

    loadStatus();
    const interval = window.setInterval(loadStatus, REFRESH_INTERVAL_MS);
    return () => {
      stopped = true;
      window.clearInterval(interval);
    };
  }, [enabled]);

  return status;
}

function useCmrPrintData(enabled) {
  const [refreshKey, setRefreshKey] = useState(0);
  const [state, setState] = useState({ loading: true, data: null, error: "" });

  useEffect(() => {
    if (!enabled) {
      setState({ loading: false, data: null, error: "" });
      return undefined;
    }

    let cancelled = false;
    setState((current) => ({ ...current, loading: true, error: "" }));
    apiJson("/api/cmrprint/data")
      .then((payload) => {
        if (!cancelled) {
          setState({ loading: false, data: payload, error: "" });
        }
      })
      .catch((loadError) => {
        if (!cancelled) {
          setState({ loading: false, data: null, error: loadError.message });
        }
      });

    return () => {
      cancelled = true;
    };
  }, [enabled, refreshKey]);

  return {
    ...state,
    refresh: () => setRefreshKey((value) => value + 1),
  };
}

function formatCmrTransportDate(dateValue = new Date()) {
  const day = String(dateValue.getDate()).padStart(2, "0");
  const month = String(dateValue.getMonth() + 1).padStart(2, "0");
  const year = String(dateValue.getFullYear());
  return `${day}-${month}-${year}`;
}

function mergeCmrFieldLines(...values) {
  const lines = [];
  for (const value of values) {
    for (const line of String(value || "")
      .split(/\n/)
      .map((part) => part.trim())
      .filter(Boolean)) {
      if (!lines.some((existing) => existing.toLowerCase() === line.toLowerCase())) {
        lines.push(line);
      }
    }
  }
  return lines.join("\n");
}

function applyCmrAssignments(target, assignments, manualFields) {
  for (const assignment of assignments || []) {
    if (!assignment?.field_name || manualFields.has(assignment.field_name)) {
      continue;
    }
    target[assignment.field_name] = assignment.value || "";
  }
}

function buildCmrCustomerBlock(customer) {
  return [customer?.name, customer?.address, customer?.city, customer?.country].filter(Boolean).join("\n");
}

function buildCmrDocumentValues(customer, exporter, transportInfo, loadingPlace, manualValues, places) {
  const values = {};
  const manualFields = new Set(["DocumentsAttached", "PackagingType", "NatureofGoods", "TransportAuthorizations"]);
  applyCmrAssignments(values, exporter?.field_assignments, manualFields);
  applyCmrAssignments(values, transportInfo?.field_assignments, manualFields);
  applyCmrAssignments(values, loadingPlace?.field_assignments, manualFields);
  applyCmrAssignments(values, customer?.field_assignments, manualFields);

  if (!String(values.ConsignorName || "").trim()) {
    values.ConsignorName = buildCmrCustomerBlock(customer);
  }

  const hasExportDate = (places || []).some((place) => place.field_name === "ExportDate");
  if (hasExportDate) {
    const dateValue = customer?.place_of_issue
      ? `${customer.place_of_issue} ${formatCmrTransportDate()}`
      : formatCmrTransportDate();
    const combinedPlaceDate = mergeCmrFieldLines(values.ConsignorRemarks, values.ExportDate, dateValue);
    values.ConsignorRemarks = combinedPlaceDate;
    values.ExportDate = combinedPlaceDate;
  }

  values.DocumentsAttached = manualValues.documentsAttached || "";
  values.PackagingType = manualValues.packagingType || "";
  values.NatureofGoods = manualValues.natureOfGoods || "xx Pal\nxx DC\nxx DCO\nxx DCS";
  values.TransportAuthorizations = manualValues.transportAuthorizations || "";
  return values;
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function buildCmrPrintHtml(title, pages) {
  const pageMarkup = pages.map((page, pageIndex) => {
    const documentWidth = page.documentWidth || CMR_A4_WIDTH;
    const documentHeight = page.documentHeight || CMR_A4_HEIGHT;
    const scaleX = CMR_A4_WIDTH / documentWidth;
    const scaleY = CMR_A4_HEIGHT / documentHeight;
    const fontScale = Math.min(scaleX, scaleY);
    return `
      <section class="cmr-print-page-sheet" data-title="${escapeHtml(page.title || title || `CMR ${pageIndex + 1}`)}">
        ${page.fields.map((field) => {
          const left = field.x * scaleX;
          const top = (field.y + (field.offset || 0)) * scaleY;
          const width = field.width * scaleX;
          const height = field.height * scaleY;
          const fontSize = Math.max(6, field.fontSize * fontScale);
          return `
            <div class="cmr-print-sheet-field" style="left:${left}px;top:${top}px;width:${width}px;min-height:${height}px;font-size:${fontSize}pt;">
              <span>${escapeHtml(field.value)}</span>
            </div>
          `;
        }).join("")}
      </section>
    `;
  }).join("");

  return `<!doctype html>
<html>
  <head>
    <meta charset="utf-8" />
    <title>${escapeHtml(title)}</title>
    <style>
      :root { color-scheme: light; }
      * { box-sizing: border-box; }
      @page { size: A4 portrait; margin: 0; }
      body { margin: 0; padding: 16px; font-family: Arial, sans-serif; background: #eef3f7; color: #10243e; }
      .cmr-print-stack { display: grid; gap: 16px; }
      .cmr-print-page-sheet {
        position: relative;
        width: 827px;
        height: 1169px;
        margin: 0 auto;
        background: white;
        overflow: hidden;
        page-break-after: always;
      }
      .cmr-print-page-sheet:last-child { page-break-after: auto; }
      .cmr-print-sheet-field {
        position: absolute;
        display: block;
        text-align: left;
        white-space: pre-wrap;
        line-height: 1.22;
        padding: 0;
        overflow: hidden;
      }
      .cmr-print-sheet-field > span {
        display: block;
        width: 100%;
        text-align: left;
        white-space: pre-wrap;
      }
      @media print {
        body { padding: 0; background: white; }
        .cmr-print-stack { gap: 0; }
        .cmr-print-page-sheet { margin: 0; }
      }
    </style>
  </head>
  <body>
    <main class="cmr-print-stack">${pageMarkup}</main>
  </body>
</html>`;
}

function openCmrPrintWindow(title, pages, autoPrint = false) {
  const popup = window.open("about:blank", "_blank", "width=1200,height=900");
  if (!popup) {
    window.alert("Allow pop-ups to open the CMR print preview.");
    return;
  }

  const html = buildCmrPrintHtml(title, pages);
  popup.document.open();
  popup.document.write(html);
  popup.document.close();
  popup.focus();

  if (autoPrint) {
    popup.addEventListener("load", () => {
      popup.focus();
      popup.print();
    }, { once: true });
  }
}

const CMR_MANUAL_FIELDS = ["DocumentsAttached", "PackagingType", "NatureofGoods", "TransportAuthorizations"];
const CMR_MENU_DEFINITIONS = [
  { key: "cmrprint", label: "CMRPrint" },
  { key: "templates", label: "Template Editor" },
  { key: "exporters", label: "Exporter Info" },
  { key: "transport", label: "Transport Info" },
  { key: "customers", label: "Customer Info" },
  { key: "loading", label: "Loading Places" },
  { key: "batch", label: "Batch Print CMRs" },
];

function sortByName(items) {
  return [...items].sort((left, right) => String(left?.name || "").localeCompare(String(right?.name || "")));
}

function blankCmrProfile() {
  return { name: "", country: "", place: "", field_assignments: [] };
}

function blankCmrCustomer() {
  return {
    name: "",
    address: "",
    city: "",
    country: "",
    vat_number: "",
    exporter_profile_name: "",
    transport_profile_name: "",
    loading_place_profile_name: "",
    place_of_issue: "",
    field_assignments: [],
  };
}

function blankCmrTemplate(name, places) {
  return {
    name: name || "",
    created_date: new Date().toISOString(),
    font_sizes: places.map((place) => ({ field_name: place.field_name, value: place.default_font_size || 9 })),
    vertical_offsets: places.map((place) => ({ field_name: place.field_name, value: 0 })),
    field_positions: places.map((place) => ({ field_name: place.field_name, x: place.default_x, y: place.default_y })),
    field_widths: places.map((place) => ({ field_name: place.field_name, value: CMR_DEFAULT_FIELD_WIDTH })),
    field_heights: places.map((place) => ({ field_name: place.field_name, value: CMR_DEFAULT_FIELD_HEIGHT })),
  };
}

function setEntryValue(entries, fieldName, value, fallback = {}) {
  const next = [...entries];
  const index = next.findIndex((entry) => entry.field_name === fieldName);
  if (index >= 0) {
    next[index] = { ...next[index], ...value };
  } else {
    next.push({ field_name: fieldName, ...fallback, ...value });
  }
  return next;
}

function getCmrDocumentBounds(places, positionMap, widthMap, heightMap) {
  const fieldBounds = places.map((place) => {
    const position = positionMap[place.field_name] || { x: place.default_x, y: place.default_y };
    const width = widthMap[place.field_name] || CMR_DEFAULT_FIELD_WIDTH;
    const height = heightMap[place.field_name] || CMR_DEFAULT_FIELD_HEIGHT;
    return { right: position.x + width, bottom: position.y + height };
  });
  const maxX = Math.max(1, ...fieldBounds.map((field) => field.right));
  const maxY = Math.max(1, ...fieldBounds.map((field) => field.bottom));
  return { width: maxX + CMR_DOCUMENT_PADDING, height: maxY + CMR_DOCUMENT_PADDING };
}

function getCmrA4Scale(documentBounds) {
  const scaleX = CMR_A4_WIDTH / Math.max(1, documentBounds?.width || CMR_A4_WIDTH);
  const scaleY = CMR_A4_HEIGHT / Math.max(1, documentBounds?.height || CMR_A4_HEIGHT);
  return { scaleX, scaleY, fontScale: Math.min(scaleX, scaleY) };
}

function CmrAssignmentsEditor({ assignments, onChange, places }) {
  const options = places.map((place) => ({ value: place.field_name, label: `${place.place_number}. ${place.description}` }));

  function updateRow(index, patch) {
    onChange(assignments.map((item, itemIndex) => (itemIndex === index ? { ...item, ...patch } : item)));
  }

  function addRow() {
    const fallback = options.find((item) => !assignments.some((entry) => entry.field_name === item.value)) || options[0];
    onChange([...assignments, { field_name: fallback?.value || "ConsignorName", value: "" }]);
  }

  return (
    <div className="cmr-assignment-editor">
      <div className="table-wrap">
        <table className="data-table compact-table">
          <thead>
            <tr><th>CMR field</th><th>Value</th><th /></tr>
          </thead>
          <tbody>
            {assignments.map((assignment, index) => (
              <tr key={`${assignment.field_name}:${index}`}>
                <td>
                  <select value={assignment.field_name} onChange={(event) => updateRow(index, { field_name: event.target.value })}>
                    {options.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}
                  </select>
                </td>
                <td>
                  <textarea rows={3} value={assignment.value || ""} onChange={(event) => updateRow(index, { value: event.target.value })} />
                </td>
                <td className="row-actions">
                  <button type="button" onClick={() => onChange(assignments.filter((_, itemIndex) => itemIndex !== index))}>Remove</button>
                </td>
              </tr>
            ))}
            {!assignments.length && (
              <tr><td colSpan="3">No field assignments yet.</td></tr>
            )}
          </tbody>
        </table>
      </div>
      <button type="button" onClick={addRow}>Add field</button>
    </div>
  );
}

function CmrProfileManager({ title, records, onChange, places }) {
  const [selectedName, setSelectedName] = useState("");
  const sortedRecords = useMemo(() => sortByName(records), [records]);

  useEffect(() => {
    if (!sortedRecords.length) {
      setSelectedName("");
      return;
    }
    if (!sortedRecords.some((item) => item.name === selectedName)) {
      setSelectedName(sortedRecords[0].name);
    }
  }, [sortedRecords, selectedName]);

  const selected = sortedRecords.find((item) => item.name === selectedName) || null;

  function replaceSelected(nextRecord) {
    if (!selected) {
      return;
    }
    onChange(sortByName(records.map((item) => (item.name === selected.name ? nextRecord : item))));
    setSelectedName(nextRecord.name || selected.name);
  }

  function addRecord() {
    const nextName = `New ${title} ${records.length + 1}`;
    const nextRecord = { ...blankCmrProfile(), name: nextName };
    onChange(sortByName([...records, nextRecord]));
    setSelectedName(nextName);
  }

  function deleteRecord() {
    if (!selected || !window.confirm(`Delete ${selected.name}?`)) {
      return;
    }
    const next = records.filter((item) => item.name !== selected.name);
    onChange(sortByName(next));
    setSelectedName(next[0]?.name || "");
  }

  return (
    <div className="cmr-editor-layout">
      <div className="cmr-editor-list data-table-card">
        <div className="section-header"><h3>{title}</h3></div>
        <div className="cmr-select-list">
          {sortedRecords.map((item) => (
            <button key={item.name} type="button" className={item.name === selectedName ? "active" : ""} onClick={() => setSelectedName(item.name)}>
              {item.name}
            </button>
          ))}
        </div>
        <div className="row-actions spread-actions">
          <button type="button" onClick={addRecord}>New</button>
          <button type="button" onClick={deleteRecord} disabled={!selected}>Delete</button>
        </div>
      </div>
      <div className="data-table-card cmr-editor-form">
        {selected ? (
          <>
            <div className="form-grid">
              <label><span>Name</span><input value={selected.name} onChange={(event) => replaceSelected({ ...selected, name: event.target.value })} /></label>
              <label><span>Country</span><input value={selected.country || ""} onChange={(event) => replaceSelected({ ...selected, country: event.target.value })} /></label>
              <label className="wide"><span>Place</span><input value={selected.place || ""} onChange={(event) => replaceSelected({ ...selected, place: event.target.value })} /></label>
            </div>
            <CmrAssignmentsEditor assignments={selected.field_assignments || []} onChange={(field_assignments) => replaceSelected({ ...selected, field_assignments })} places={places} />
          </>
        ) : (
          <div className="notice">No records yet.</div>
        )}
      </div>
    </div>
  );
}

function CmrCustomerManager({ customers, exporters, transportInfos, loadingPlaces, onChange, places }) {
  const [selectedName, setSelectedName] = useState("");
  const sortedCustomers = useMemo(() => sortByName(customers), [customers]);

  useEffect(() => {
    if (!sortedCustomers.length) {
      setSelectedName("");
      return;
    }
    if (!sortedCustomers.some((item) => item.name === selectedName)) {
      setSelectedName(sortedCustomers[0].name);
    }
  }, [sortedCustomers, selectedName]);

  const selected = sortedCustomers.find((item) => item.name === selectedName) || null;

  function replaceSelected(nextRecord) {
    if (!selected) {
      return;
    }
    onChange(sortByName(customers.map((item) => (item.name === selected.name ? nextRecord : item))));
    setSelectedName(nextRecord.name || selected.name);
  }

  function addCustomer() {
    const nextName = `New customer ${customers.length + 1}`;
    const nextRecord = { ...blankCmrCustomer(), name: nextName };
    onChange(sortByName([...customers, nextRecord]));
    setSelectedName(nextName);
  }

  function deleteCustomer() {
    if (!selected || !window.confirm(`Delete ${selected.name}?`)) {
      return;
    }
    const next = customers.filter((item) => item.name !== selected.name);
    onChange(sortByName(next));
    setSelectedName(next[0]?.name || "");
  }

  return (
    <div className="cmr-editor-layout">
      <div className="cmr-editor-list data-table-card">
        <div className="section-header"><h3>Customer Info</h3></div>
        <div className="cmr-select-list">
          {sortedCustomers.map((item) => (
            <button key={item.name} type="button" className={item.name === selectedName ? "active" : ""} onClick={() => setSelectedName(item.name)}>
              {item.name}
            </button>
          ))}
        </div>
        <div className="row-actions spread-actions">
          <button type="button" onClick={addCustomer}>New</button>
          <button type="button" onClick={deleteCustomer} disabled={!selected}>Delete</button>
        </div>
      </div>
      <div className="data-table-card cmr-editor-form">
        {selected ? (
          <>
            <div className="form-grid">
              <label className="wide"><span>Name</span><input value={selected.name} onChange={(event) => replaceSelected({ ...selected, name: event.target.value })} /></label>
              <label className="wide"><span>Address</span><textarea rows={3} value={selected.address || ""} onChange={(event) => replaceSelected({ ...selected, address: event.target.value })} /></label>
              <label><span>City</span><input value={selected.city || ""} onChange={(event) => replaceSelected({ ...selected, city: event.target.value })} /></label>
              <label><span>Country</span><input value={selected.country || ""} onChange={(event) => replaceSelected({ ...selected, country: event.target.value })} /></label>
              <label><span>VAT</span><input value={selected.vat_number || ""} onChange={(event) => replaceSelected({ ...selected, vat_number: event.target.value })} /></label>
              <label><span>Place/date</span><input value={selected.place_of_issue || ""} onChange={(event) => replaceSelected({ ...selected, place_of_issue: event.target.value })} /></label>
              <label><span>Exporter</span><select value={selected.exporter_profile_name || ""} onChange={(event) => replaceSelected({ ...selected, exporter_profile_name: event.target.value })}><option value="">-</option>{sortByName(exporters).map((item) => <option key={item.name} value={item.name}>{item.name}</option>)}</select></label>
              <label><span>Transport</span><select value={selected.transport_profile_name || ""} onChange={(event) => replaceSelected({ ...selected, transport_profile_name: event.target.value })}><option value="">-</option>{sortByName(transportInfos).map((item) => <option key={item.name} value={item.name}>{item.name}</option>)}</select></label>
              <label className="wide"><span>Loading place</span><select value={selected.loading_place_profile_name || ""} onChange={(event) => replaceSelected({ ...selected, loading_place_profile_name: event.target.value })}><option value="">-</option>{sortByName(loadingPlaces).map((item) => <option key={item.name} value={item.name}>{item.name}</option>)}</select></label>
            </div>
            <div className="sidebar-note">Manual fields stay free: 5, 7, 9, 17.</div>
            <CmrAssignmentsEditor assignments={selected.field_assignments || []} onChange={(field_assignments) => replaceSelected({ ...selected, field_assignments })} places={places.filter((place) => !CMR_MANUAL_FIELDS.includes(place.field_name))} />
          </>
        ) : (
          <div className="notice">No customers yet.</div>
        )}
      </div>
    </div>
  );
}

function CmrTemplateEditor({ templates, places, defaultTemplateName, onSaveTemplate, onDeleteTemplate }) {
  const [selectedTemplateName, setSelectedTemplateName] = useState("");
  const [draft, setDraft] = useState(null);
  const [status, setStatus] = useState({ error: "", message: "", busy: false });
  const sortedTemplates = useMemo(() => sortByName(templates), [templates]);

  useEffect(() => {
    const firstName = defaultTemplateName || sortedTemplates[0]?.name || "Draft template";
    if (!selectedTemplateName) {
      setSelectedTemplateName(firstName);
    }
  }, [defaultTemplateName, sortedTemplates, selectedTemplateName]);

  useEffect(() => {
    const current = sortedTemplates.find((item) => item.name === selectedTemplateName) || blankCmrTemplate(selectedTemplateName, places);
    setDraft(JSON.parse(JSON.stringify(current)));
  }, [selectedTemplateName, sortedTemplates, places]);

  const positionMap = useMemo(() => Object.fromEntries((draft?.field_positions || []).map((entry) => [entry.field_name, entry])), [draft]);
  const widthMap = useMemo(() => Object.fromEntries((draft?.field_widths || []).map((entry) => [entry.field_name, entry.value])), [draft]);
  const heightMap = useMemo(() => Object.fromEntries((draft?.field_heights || []).map((entry) => [entry.field_name, entry.value])), [draft]);
  const fontMap = useMemo(() => Object.fromEntries((draft?.font_sizes || []).map((entry) => [entry.field_name, entry.value])), [draft]);
  const offsetMap = useMemo(() => Object.fromEntries((draft?.vertical_offsets || []).map((entry) => [entry.field_name, entry.value])), [draft]);
  const previewBounds = useMemo(() => getCmrDocumentBounds(places, positionMap, widthMap, heightMap), [places, positionMap, widthMap, heightMap]);
  const previewScale = useMemo(() => getCmrA4Scale(previewBounds), [previewBounds]);

  function updateField(fieldName, key, value) {
    if (!draft) {
      return;
    }
    if (key === "font") {
      setDraft({ ...draft, font_sizes: setEntryValue(draft.font_sizes || [], fieldName, { value: Number(value) || 0 }) });
      return;
    }
    if (key === "offset") {
      setDraft({ ...draft, vertical_offsets: setEntryValue(draft.vertical_offsets || [], fieldName, { value: Number(value) || 0 }) });
      return;
    }
    if (key === "width") {
      setDraft({ ...draft, field_widths: setEntryValue(draft.field_widths || [], fieldName, { value: Number(value) || 0 }) });
      return;
    }
    if (key === "height") {
      setDraft({ ...draft, field_heights: setEntryValue(draft.field_heights || [], fieldName, { value: Number(value) || 0 }) });
      return;
    }
    if (key === "x") {
      setDraft({ ...draft, field_positions: setEntryValue(draft.field_positions || [], fieldName, { x: Number(value) || 0 }, { y: positionMap[fieldName]?.y || 0 }) });
      return;
    }
    if (key === "y") {
      setDraft({ ...draft, field_positions: setEntryValue(draft.field_positions || [], fieldName, { y: Number(value) || 0 }, { x: positionMap[fieldName]?.x || 0 }) });
    }
  }

  async function saveTemplate() {
    if (!draft?.name?.trim()) {
      setStatus({ busy: false, error: "Template name is required", message: "" });
      return;
    }
    setStatus({ busy: true, error: "", message: "" });
    try {
      await onSaveTemplate(draft);
      setStatus({ busy: false, error: "", message: "Template saved" });
    } catch (error) {
      setStatus({ busy: false, error: error.message, message: "" });
    }
  }

  async function deleteTemplate() {
    if (!draft?.name || !window.confirm(`Delete template ${draft.name}?`)) {
      return;
    }
    setStatus({ busy: true, error: "", message: "" });
    try {
      await onDeleteTemplate(draft.name);
      setSelectedTemplateName("");
      setStatus({ busy: false, error: "", message: "Template deleted" });
    } catch (error) {
      setStatus({ busy: false, error: error.message, message: "" });
    }
  }

  function newTemplate() {
    const nextName = `template-${Date.now()}`;
    setSelectedTemplateName(nextName);
    setDraft(blankCmrTemplate(nextName, places));
  }

  return (
    <div className="data-table-card cmr-template-editor">
      <div className="cmr-template-topbar">
        <label><span>Template</span><select value={selectedTemplateName} onChange={(event) => setSelectedTemplateName(event.target.value)}><option value="">Choose template</option>{sortedTemplates.map((item) => <option key={item.name} value={item.name}>{item.name}</option>)}</select></label>
        <div className="row-actions spread-actions">
          <button type="button" onClick={newTemplate}>New</button>
          <button type="button" onClick={saveTemplate} disabled={status.busy}>{status.busy ? "Saving..." : "Save"}</button>
          <button type="button" onClick={deleteTemplate} disabled={status.busy || !draft?.name}>Delete</button>
        </div>
      </div>
      {status.message && <div className="notice">{status.message}</div>}
      {status.error && <div className="notice danger">{status.error}</div>}
      {draft && (
        <>
          <label><span>Template name</span><input value={draft.name} onChange={(event) => setDraft({ ...draft, name: event.target.value })} /></label>
          <div className="cmr-template-grid">
            <div className="cmr-preview-shell">
              <div className="cmr-preview-board cmr-a4-preview-board" style={{ width: CMR_A4_WIDTH, height: CMR_A4_HEIGHT }}>
                {places.map((place) => {
                  const position = positionMap[place.field_name] || { x: place.default_x, y: place.default_y };
                  const width = widthMap[place.field_name] || CMR_DEFAULT_FIELD_WIDTH;
                  const height = heightMap[place.field_name] || CMR_DEFAULT_FIELD_HEIGHT;
                  const offset = offsetMap[place.field_name] || 0;
                  const fontSize = fontMap[place.field_name] || place.default_font_size || 9;
                  return (
                    <div key={place.field_name} className="cmr-preview-field" style={{ left: position.x * previewScale.scaleX, top: (position.y + offset) * previewScale.scaleY, width: width * previewScale.scaleX, minHeight: height * previewScale.scaleY }}>
                      <strong>{place.place_number}</strong>
                      <span style={{ fontSize: Math.max(6, fontSize * previewScale.fontScale) }}>{place.description}</span>
                    </div>
                  );
                })}
              </div>
            </div>
            <div className="table-wrap cmr-template-settings">
              <table className="data-table compact-table">
                <thead>
                  <tr><th>Place</th><th>Font</th><th>Offset</th><th>W</th><th>H</th><th>X</th><th>Y</th></tr>
                </thead>
                <tbody>
                  {places.map((place) => (
                    <tr key={place.field_name}>
                      <td>{place.place_number}. {place.description}</td>
                      <td><input type="number" value={fontMap[place.field_name] || place.default_font_size || 9} onChange={(event) => updateField(place.field_name, "font", event.target.value)} /></td>
                      <td><input type="number" value={offsetMap[place.field_name] || 0} onChange={(event) => updateField(place.field_name, "offset", event.target.value)} /></td>
                      <td><input type="number" value={widthMap[place.field_name] || CMR_DEFAULT_FIELD_WIDTH} onChange={(event) => updateField(place.field_name, "width", event.target.value)} /></td>
                      <td><input type="number" value={heightMap[place.field_name] || CMR_DEFAULT_FIELD_HEIGHT} onChange={(event) => updateField(place.field_name, "height", event.target.value)} /></td>
                      <td><input type="number" value={positionMap[place.field_name]?.x ?? place.default_x} onChange={(event) => updateField(place.field_name, "x", event.target.value)} /></td>
                      <td><input type="number" value={positionMap[place.field_name]?.y ?? place.default_y} onChange={(event) => updateField(place.field_name, "y", event.target.value)} /></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </>
      )}
    </div>
  );
}

function CmrBatchPrintView({ customers, buildPrintPage, defaultManualValues }) {
  const [search, setSearch] = useState("");
  const [selectedNames, setSelectedNames] = useState([]);
  const [overrides, setOverrides] = useState({});
  const visibleCustomers = useMemo(() => sortByName(customers).filter((item) => `${item.name} ${item.address || ""}`.toLowerCase().includes(search.trim().toLowerCase())), [customers, search]);

  function customerManualValues(name) {
    const manualOverride = overrides[name] || {};
    return {
      documentsAttached: manualOverride.documentsAttached ?? defaultManualValues.documentsAttached,
      packagingType: manualOverride.packagingType ?? defaultManualValues.packagingType,
      natureOfGoods: manualOverride.natureOfGoods ?? defaultManualValues.natureOfGoods,
      transportAuthorizations: manualOverride.transportAuthorizations ?? defaultManualValues.transportAuthorizations,
    };
  }

  function updateCustomerManualValue(name, field, value) {
    setOverrides((current) => ({
      ...current,
      [name]: {
        ...current[name],
        [field]: value,
      },
    }));
  }

  function toggle(name) {
    setSelectedNames((current) => current.includes(name) ? current.filter((item) => item !== name) : [...current, name]);
  }

  function openBatch(autoPrint) {
    const pages = selectedNames
      .map((name) => visibleCustomers.find((item) => item.name === name) || customers.find((item) => item.name === name))
      .filter(Boolean)
      .map((customer) => buildPrintPage(customer, customerManualValues(customer.name)));
    if (!pages.length) {
      window.alert("Choose at least one customer.");
      return;
    }
    openCmrPrintWindow("Batch CMR print", pages, autoPrint);
  }

  return (
    <div className="data-table-card cmr-batch-panel">
      <div className="cmr-batch-header">
        <label className="wide"><span>Search customer</span><input value={search} onChange={(event) => setSearch(event.target.value)} /></label>
        <div className="row-actions spread-actions">
          <button type="button" onClick={() => openBatch(false)}>Preview selected</button>
          <button type="button" className="primary" onClick={() => openBatch(true)}>Print selected</button>
        </div>
      </div>
      <div className="cmr-batch-list-table">
        {visibleCustomers.map((customer) => {
          const manualOverride = customerManualValues(customer.name);
          return (
            <div key={customer.name} className="cmr-batch-row">
              <div>
                <strong>{customer.name}</strong>
                <div className="sidebar-note">{customer.address || "-"}</div>
              </div>
              <label className="cmr-batch-field cmr-batch-field-5">
                <span>5 - Documents attached</span>
                <textarea rows={3} value={manualOverride.documentsAttached} onChange={(event) => updateCustomerManualValue(customer.name, "documentsAttached", event.target.value)} />
              </label>
              <label className="cmr-batch-field cmr-batch-field-7">
                <span>7 - Packaging / marks</span>
                <textarea rows={3} value={manualOverride.packagingType} onChange={(event) => updateCustomerManualValue(customer.name, "packagingType", event.target.value)} />
              </label>
              <label className="cmr-batch-field cmr-batch-field-9">
                <span>9 - Nature of goods</span>
                <textarea rows={3} value={manualOverride.natureOfGoods} onChange={(event) => updateCustomerManualValue(customer.name, "natureOfGoods", event.target.value)} />
              </label>
              <label className="cmr-batch-field cmr-batch-field-17">
                <span>17 - Transport authorizations</span>
                <textarea rows={3} value={manualOverride.transportAuthorizations} onChange={(event) => updateCustomerManualValue(customer.name, "transportAuthorizations", event.target.value)} />
              </label>
              <label className="checkbox-row">
                <input type="checkbox" checked={selectedNames.includes(customer.name)} onChange={() => toggle(customer.name)} />
                <span>Add</span>
              </label>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function CmrPrintPage({ currentUser }) {
  const enabled = hasPermission(currentUser, PERMISSIONS.CMR_VIEW);
  const { loading, data, error, refresh } = useCmrPrintData(enabled);
  const [activeMenu, setActiveMenu] = useState("cmrprint");
  const [message, setMessage] = useState("");
  const [saveError, setSaveError] = useState("");
  const [saving, setSaving] = useState(false);
  const [draftData, setDraftData] = useState(null);
  const [hasUnsavedChanges, setHasUnsavedChanges] = useState(false);
  const [selectedCustomerName, setSelectedCustomerName] = useState("");
  const [selectedTemplateNameOverride, setSelectedTemplateNameOverride] = useState("");
  const [manualValues, setManualValues] = useState({
    documentsAttached: "",
    packagingType: "",
    natureOfGoods: "xx Pal\nxx DC\nxx DCO\nxx DCS",
    transportAuthorizations: "",
  });

  useEffect(() => {
    if (!data) {
      return;
    }
    setDraftData({
      available: data.available,
      customers: data.customers || [],
      exporters: data.exporters || [],
      transport_infos: data.transport_infos || [],
      loading_places: data.loading_places || [],
      templates: data.templates || [],
      settings: data.settings || {},
      places: data.places || [],
      can_manage: data.can_manage,
      data_dir: data.data_dir,
      templates_dir: data.templates_dir,
      debug_candidates: data.debug_candidates || [],
    });
    setHasUnsavedChanges(false);
    setSaveError("");
  }, [data]);

  const cmrData = draftData || data;
  const places = cmrData?.places || [];
  const customers = cmrData?.customers || [];
  const exporters = cmrData?.exporters || [];
  const transportInfos = cmrData?.transport_infos || [];
  const loadingPlaces = cmrData?.loading_places || [];
  const templates = cmrData?.templates || [];
  const canManage = Boolean(cmrData?.can_manage);
  const latestCmrCollectionsRef = useRef({
    customers,
    exporters,
    transport_infos: transportInfos,
    loading_places: loadingPlaces,
  });
  const visibleMenus = canManage ? CMR_MENU_DEFINITIONS : CMR_MENU_DEFINITIONS.filter((item) => item.key === "cmrprint");
  const settings = cmrData?.settings || {};

  latestCmrCollectionsRef.current = {
    customers,
    exporters,
    transport_infos: transportInfos,
    loading_places: loadingPlaces,
  };

  useEffect(() => {
    if (!visibleMenus.some((item) => item.key === activeMenu)) {
      setActiveMenu("cmrprint");
    }
  }, [visibleMenus, activeMenu]);

  useEffect(() => {
    if (!customers.length) {
      setSelectedCustomerName("");
      return;
    }
    if (!customers.some((item) => item.name === selectedCustomerName)) {
      setSelectedCustomerName(customers[0].name);
    }
  }, [customers, selectedCustomerName]);

  useEffect(() => {
    const defaultTemplateName = settings.cmr_default_template_name || templates[0]?.name || "";
    if (!selectedTemplateNameOverride) {
      setSelectedTemplateNameOverride(defaultTemplateName);
      return;
    }
    if (!templates.some((item) => item.name === selectedTemplateNameOverride)) {
      setSelectedTemplateNameOverride(defaultTemplateName);
    }
  }, [settings.cmr_default_template_name, templates, selectedTemplateNameOverride]);

  const selectedTemplateName = selectedTemplateNameOverride || settings.cmr_default_template_name || templates[0]?.name || "";
  const template = templates.find((item) => item.name === selectedTemplateName) || templates[0] || null;
  const customer = customers.find((item) => item.name === selectedCustomerName) || null;
  const exporter = exporters.find((item) => item.name === customer?.exporter_profile_name) || null;
  const transportInfo = transportInfos.find((item) => item.name === customer?.transport_profile_name) || null;
  const loadingPlace = loadingPlaces.find((item) => item.name === customer?.loading_place_profile_name) || null;

  const documentValues = useMemo(
    () => buildCmrDocumentValues(customer, exporter, transportInfo, loadingPlace, manualValues, places),
    [customer, exporter, transportInfo, loadingPlace, manualValues, places],
  );
  const positionMap = useMemo(() => Object.fromEntries((template?.field_positions || []).map((entry) => [entry.field_name, entry])), [template]);
  const widthMap = useMemo(() => Object.fromEntries((template?.field_widths || []).map((entry) => [entry.field_name, entry.value])), [template]);
  const heightMap = useMemo(() => Object.fromEntries((template?.field_heights || []).map((entry) => [entry.field_name, entry.value])), [template]);
  const fontMap = useMemo(() => Object.fromEntries((template?.font_sizes || []).map((entry) => [entry.field_name, entry.value])), [template]);
  const offsetMap = useMemo(() => Object.fromEntries((template?.vertical_offsets || []).map((entry) => [entry.field_name, entry.value])), [template]);
  const previewBounds = useMemo(() => getCmrDocumentBounds(places, positionMap, widthMap, heightMap), [places, positionMap, widthMap, heightMap]);
  const previewScale = useMemo(() => getCmrA4Scale(previewBounds), [previewBounds]);

  const autofillSummary = useMemo(() => {
    if (!customer) {
      return "Select a customer to load linked profiles and CMR field values.";
    }
    const lines = [
      `Customer: ${customer.name}`,
      `Template: ${selectedTemplateName || "-"}`,
      `Field 5: ${documentValues.DocumentsAttached || "-"}`,
      `Field 21: ${documentValues.ExportDate || "-"}`,
    ];
    for (const place of places) {
      const value = String(documentValues[place.field_name] || "").trim();
      if (!value || CMR_MANUAL_FIELDS.includes(place.field_name)) {
        continue;
      }
      lines.push(`${place.description}: ${value.replace(/\n/g, " | ")}`);
    }
    return lines.join("\n");
  }, [customer, documentValues, places, selectedTemplateName]);

  function updateCollections(patch) {
    setDraftData((current) => ({ ...current, ...patch }));
    setHasUnsavedChanges(true);
  }

  async function saveAppData(patchMessage) {
    if (!cmrData) {
      return;
    }
    if (document.activeElement instanceof HTMLElement) {
      document.activeElement.blur();
    }
    setSaving(true);
    setMessage("");
    setSaveError("");
    try {
      await new Promise((resolve) => window.requestAnimationFrame(resolve));
      const latestCollections = latestCmrCollectionsRef.current;
      const payload = await apiJson("/api/cmrprint/app-data", {
        method: "PATCH",
        body: JSON.stringify(latestCollections),
      });
      setDraftData({ ...payload, settings: data?.settings || cmrData.settings, can_manage: cmrData.can_manage });
      setMessage(patchMessage || "CMR data saved.");
      setHasUnsavedChanges(false);
      refresh();
    } catch (saveError) {
      setSaveError(saveError.message);
    } finally {
      setSaving(false);
    }
  }

  useEffect(() => {
    if (!hasUnsavedChanges) {
      return undefined;
    }
    const handleBeforeUnload = (event) => {
      event.preventDefault();
      event.returnValue = "";
      return "";
    };
    window.addEventListener("beforeunload", handleBeforeUnload);
    return () => window.removeEventListener("beforeunload", handleBeforeUnload);
  }, [hasUnsavedChanges]);

  async function saveTemplate(templatePayload) {
    const payload = await apiJson("/api/cmrprint/template", { method: "PUT", body: JSON.stringify({ template: templatePayload }) });
    setDraftData((current) => ({ ...current, templates: payload.templates || [] }));
    setSelectedTemplateNameOverride(templatePayload?.name || "");
    refresh();
  }

  async function deleteTemplate(templateName) {
    const payload = await apiJson(`/api/cmrprint/template/${encodeURIComponent(templateName)}`, { method: "DELETE" });
    setDraftData((current) => ({ ...current, templates: payload.templates || [] }));
    refresh();
  }

  async function saveDefaultTemplate() {
    const templateName = selectedTemplateNameOverride || "";
    if (!templateName) {
      setSaveError("Choose a template first.");
      return;
    }
    setSaving(true);
    setMessage("");
    setSaveError("");
    try {
      const payload = await apiJson("/api/cmrprint/settings", {
        method: "PATCH",
        body: JSON.stringify({ cmr_default_template_name: templateName }),
      });
      setDraftData((current) => ({
        ...current,
        settings: {
          ...(current?.settings || {}),
          ...(payload.settings || {}),
        },
      }));
      setMessage(`Default CMR template saved as ${templateName}.`);
      refresh();
    } catch (error) {
      setSaveError(error.message);
    } finally {
      setSaving(false);
    }
  }

  function buildPrintPageForCustomer(customerRecord, manualOverride = {}) {
    if (!customerRecord) {
      return null;
    }
    const pageExporter = exporters.find((item) => item.name === customerRecord.exporter_profile_name) || null;
    const pageTransport = transportInfos.find((item) => item.name === customerRecord.transport_profile_name) || null;
    const pageLoading = loadingPlaces.find((item) => item.name === customerRecord.loading_place_profile_name) || null;
    const values = buildCmrDocumentValues(customerRecord, pageExporter, pageTransport, pageLoading, { ...manualValues, ...manualOverride }, places);
    const fields = places.map((place) => ({
      x: positionMap[place.field_name]?.x ?? place.default_x,
      y: positionMap[place.field_name]?.y ?? place.default_y,
      width: widthMap[place.field_name] || CMR_DEFAULT_FIELD_WIDTH,
      height: heightMap[place.field_name] || CMR_DEFAULT_FIELD_HEIGHT,
      fontSize: fontMap[place.field_name] || place.default_font_size || 9,
      offset: offsetMap[place.field_name] || 0,
      value: String(values[place.field_name] || "").trim(),
    })).filter((field) => field.value);
    return {
      title: customerRecord.name,
      subtitle: `${selectedTemplateName || "CMR template"} | ${customerRecord.country || ""}`,
      documentWidth: previewBounds.width,
      documentHeight: previewBounds.height,
      fields,
    };
  }

  function openCurrentCustomerPrint(autoPrint) {
    const page = buildPrintPageForCustomer(customer);
    if (!page) {
      window.alert("Choose a customer first.");
      return;
    }
    openCmrPrintWindow(`${customer.name} CMR`, [page], autoPrint);
  }

  if (loading) {
    return <div className="notice">Loading CMR Print data...</div>;
  }
  if (error) {
    return <div className="notice danger">Unable to load CMR Print data: {error}</div>;
  }
  if (!cmrData?.available) {
    return (
      <div className="data-table-card cmr-missing-data">
        <div className="notice danger">No CMR Print data folder was found yet.</div>
        {Array.isArray(cmrData?.debug_candidates) && cmrData.debug_candidates.length > 0 && (
          <div className="table-wrap">
            <table className="data-table compact-table">
              <thead>
                <tr><th>Checked path</th><th>Folder</th><th>app-data.xml</th><th>Templates</th></tr>
              </thead>
              <tbody>
                {cmrData.debug_candidates.map((item) => (
                  <tr key={item.path}>
                    <td>{item.path}</td>
                    <td>{item.exists ? "yes" : "no"}</td>
                    <td>{item.has_app_data ? "yes" : "no"}</td>
                    <td>{item.has_templates_dir ? "yes" : "no"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    );
  }

  return (
    <section className="overview-stack cmr-print-page">
      <div className="tab-strip cmr-submenu-strip">
        {visibleMenus.map((menu) => (
          <button
            key={menu.key}
            type="button"
            className={activeMenu === menu.key ? "active" : ""}
            onClick={() => {
              if (menu.key !== activeMenu && hasUnsavedChanges && !window.confirm("You have unsaved CMR changes. Leave this menu without saving?")) {
                return;
              }
              setActiveMenu(menu.key);
            }}
          >
            {menu.label}
          </button>
        ))}
      </div>
      {message && <div className="notice">{message}</div>}
      {saveError && <div className="notice danger">{saveError}</div>}
      {hasUnsavedChanges && canManage && ["customers", "exporters", "transport", "loading"].includes(activeMenu) && (
        <div className="notice danger">You have unsaved CMR changes. Click Save before leaving this page.</div>
      )}

      {activeMenu === "cmrprint" && (
        <div className="data-table-card cmr-print-grid">
          <div className="cmr-print-meta-row">
            <div className="notice">Active template: <strong>{selectedTemplateName || "No template"}</strong></div>
          </div>
          <div className="form-grid cmr-print-selectors">
            <label><span>Customer</span><select value={selectedCustomerName} onChange={(event) => setSelectedCustomerName(event.target.value)}><option value="">Choose customer</option>{sortByName(customers).map((item) => <option key={item.name} value={item.name}>{item.name}</option>)}</select></label>
            <label><span>Template</span><select value={selectedTemplateNameOverride} onChange={(event) => setSelectedTemplateNameOverride(event.target.value)}><option value="">Choose template</option>{sortByName(templates).map((item) => <option key={item.name} value={item.name}>{item.name}</option>)}</select></label>
            <label><span>Exporter</span><input value={exporter?.name || "-"} readOnly /></label>
            <label><span>Transport</span><input value={transportInfo?.name || "-"} readOnly /></label>
            <label className="wide"><span>Customer block</span><textarea rows={4} value={buildCmrCustomerBlock(customer)} readOnly /></label>
          </div>
          <div className="cmr-print-toolbar"><div className="row-actions spread-actions"><button type="button" onClick={() => openCurrentCustomerPrint(false)}>Preview Print</button><button type="button" className="primary" onClick={() => openCurrentCustomerPrint(true)}>Print CMR</button>{canManage && <button type="button" onClick={saveDefaultTemplate} disabled={saving || !selectedTemplateNameOverride}>{saving ? "Saving..." : "Save as default template"}</button>}</div></div>
          <div className="cmr-print-workspace">
            <div className="cmr-print-fields">
              <label><span>Field 5 - Documents attached</span><textarea rows={4} value={manualValues.documentsAttached} onChange={(event) => setManualValues({ ...manualValues, documentsAttached: event.target.value })} /></label>
              <label><span>Field 7 - Packaging / marks</span><textarea rows={4} value={manualValues.packagingType} onChange={(event) => setManualValues({ ...manualValues, packagingType: event.target.value })} /></label>
              <label><span>Field 9 - Nature of goods</span><textarea rows={5} value={manualValues.natureOfGoods} onChange={(event) => setManualValues({ ...manualValues, natureOfGoods: event.target.value })} /></label>
              <label><span>Field 17 - Transport authorizations</span><textarea rows={4} value={manualValues.transportAuthorizations} onChange={(event) => setManualValues({ ...manualValues, transportAuthorizations: event.target.value })} /></label>
              <label><span>Autofill summary</span><textarea rows={12} value={autofillSummary} readOnly /></label>
            </div>
            <div className="cmr-preview-shell">
              <div className="cmr-preview-board cmr-a4-preview-board" style={{ width: CMR_A4_WIDTH, height: CMR_A4_HEIGHT }}>
                {places.map((place) => {
                  const position = positionMap[place.field_name] || { x: place.default_x, y: place.default_y };
                  const width = widthMap[place.field_name] || CMR_DEFAULT_FIELD_WIDTH;
                  const height = heightMap[place.field_name] || CMR_DEFAULT_FIELD_HEIGHT;
                  const offset = offsetMap[place.field_name] || 0;
                  const fontSize = fontMap[place.field_name] || place.default_font_size || 9;
                  const value = String(documentValues[place.field_name] || "").trim();
                  return <div key={place.field_name} className="cmr-preview-field" style={{ left: position.x * previewScale.scaleX, top: (position.y + offset) * previewScale.scaleY, width: width * previewScale.scaleX, minHeight: height * previewScale.scaleY }}><strong>{place.place_number}</strong><span style={{ fontSize: Math.max(6, fontSize * previewScale.fontScale) }}>{value || place.field_name}</span></div>;
                })}
              </div>
            </div>
          </div>
        </div>
      )}

      {canManage && activeMenu === "templates" && <CmrTemplateEditor templates={templates} places={places} defaultTemplateName={selectedTemplateName} onSaveTemplate={saveTemplate} onDeleteTemplate={deleteTemplate} />}
      {canManage && activeMenu === "exporters" && <><CmrProfileManager title="Exporter" records={exporters} onChange={(value) => updateCollections({ exporters: value })} places={places} /><div className="row-actions spread-actions"><button type="button" className="primary" onClick={() => saveAppData("Exporter info saved.")} disabled={saving}>{saving ? "Saving..." : "Save"}</button></div></>}
      {canManage && activeMenu === "transport" && <><CmrProfileManager title="Transport" records={transportInfos} onChange={(value) => updateCollections({ transport_infos: value })} places={places} /><div className="row-actions spread-actions"><button type="button" className="primary" onClick={() => saveAppData("Transport info saved.")} disabled={saving}>{saving ? "Saving..." : "Save"}</button></div></>}
      {canManage && activeMenu === "loading" && <><CmrProfileManager title="Loading place" records={loadingPlaces} onChange={(value) => updateCollections({ loading_places: value })} places={places} /><div className="row-actions spread-actions"><button type="button" className="primary" onClick={() => saveAppData("Loading places saved.")} disabled={saving}>{saving ? "Saving..." : "Save"}</button></div></>}
      {canManage && activeMenu === "customers" && <><CmrCustomerManager customers={customers} exporters={exporters} transportInfos={transportInfos} loadingPlaces={loadingPlaces} onChange={(value) => updateCollections({ customers: value })} places={places} /><div className="row-actions spread-actions"><button type="button" className="primary" onClick={() => saveAppData("Customer info saved.")} disabled={saving}>{saving ? "Saving..." : "Save"}</button></div></>}
      {canManage && activeMenu === "batch" && <CmrBatchPrintView customers={customers} buildPrintPage={buildPrintPageForCustomer} defaultManualValues={manualValues} />}
    </section>
  );
}

const UKDOCS_CATEGORY_DEFINITIONS = [
  { code: "508", label: "Flowers", shortLabel: "Flowers" },
  { code: "515", label: "Accessories", shortLabel: "Accessories" },
  { code: "1000", label: "Bouquets", shortLabel: "Bouquets" },
  { code: "920", label: "Plants", shortLabel: "Plants" },
];

const UKDOCS_COMPANY_FIELDS = [
  ["company_name", "Company name"],
  ["address", "Address", "textarea"],
  ["phone", "Phone"],
  ["email", "Email"],
  ["website", "Website"],
  ["vat_number", "VAT number"],
  ["eori_number", "EORI number"],
  ["chamber_of_commerce_number", "Chamber of Commerce number"],
  ["iban", "IBAN"],
  ["bic_swift", "BIC / SWIFT"],
  ["rex_registration", "REX registration"],
  ["default_footer_text", "Default footer text", "textarea"],
  ["preferential_origin_declaration", "Preferential origin declaration", "textarea"],
  ["logo_name", "Logo file name"],
];

const UKDOCS_CUSTOMER_FIELDS = [
  ["customer_name", "Customer name"],
  ["match_hub_code", "Hub code match", "textarea"],
  ["match_remark", "Remark match", "textarea"],
  ["customer_address", "Customer address", "textarea"],
  ["vat_number", "VAT number"],
  ["eori_number", "EORI number"],
  ["importer_number", "Importer number"],
  ["default_owner", "Default owner (export header)"],
  ["default_importer", "Default importer number / reference"],
  ["default_delivery_terms", "Default delivery terms"],
  ["default_city", "Default city"],
  ["default_uk_arrival_port", "Default UK arrival port"],
  ["default_currency", "Default currency"],
  ["ready_email_subject", "Ready email subject template", "textarea"],
  ["ready_email_body", "Ready email body template", "textarea"],
  ["default_invoice_language_text", "Default invoice language / text", "textarea"],
  ["default_document_references", "Default document references", "textarea"],
];

const UKDOCS_CUSTOMER_INVOICE_VISIBILITY_FIELDS = [
  ["show_invoice_vat_number", "Show VAT number on invoice"],
  ["show_invoice_eori_number", "Show EORI number on invoice"],
  ["show_invoice_importer_number", "Show importer / DAN number on invoice"],
];

const UKDOCS_CUSTOMER_REQUIRED_DOCUMENT_FIELDS = [
  ["required_phyto", "Require phytosanitary files"],
  ["required_export_extra", "Require second export file"],
  ["required_generated_export", "Require generated export workbook"],
  ["required_generated_invoices", "Require generated invoice workbooks"],
];

const UKDOCS_CUSTOMER_MENU_DOCUMENT_FIELDS = [
  ["menu_show_ukdocsinspection_inspection_list", "Phyto inspection - Inspection list"],
  ["menu_show_ukdocsinspection_locations_file", "Phyto inspection - Locations file"],
  ["menu_show_ukdocsinspection_phyto", "Phyto inspection - Phytosanitary document"],
  ["menu_show_ukdocsinspection_export_extra", "Phyto inspection - Second export file"],
  ["menu_show_ukdocsinspection_generated_invoices", "Phyto inspection - Invoices generate"],
  ["menu_show_ukdocsinspection_generated_export", "Phyto inspection - Export file generated"],
  ["menu_show_ukdocsprint_phyto", "UKdocs Print - Phytosanitary document"],
  ["menu_show_ukdocsprint_export_extra", "UKdocs Print - Second export file"],
  ["menu_show_ukdocsprint_generated_invoices", "UKdocs Print - Invoices generate"],
  ["menu_show_ukdocsprint_generated_export", "UKdocs Print - Export file generated"],
  ["menu_show_ukdocsprint_inspection_list", "UKdocs Print - Inspection list"],
  ["menu_show_ukdocsprint_locations_file", "UKdocs Print - Locations file"],
];

const UKDOCS_EXPORT_DEFAULT_FIELDS = [
  ["destination_country", "Country of destination"],
  ["regulation", "Regulation"],
  ["border_transport_mode", "Border transport mode"],
  ["border_transport_nationality", "Border transport nationality"],
  ["customs_office_of_exit", "Customs office of exit"],
  ["location", "Location"],
  ["delivery_terms", "Delivery terms"],
  ["delivery_terms_city", "Delivery terms city"],
  ["currency", "Currency"],
  ["freight_costs", "Freight costs"],
  ["insurance", "Insurance"],
  ["importer_field", "Importer field", "textarea"],
  ["vessel_field", "Vessel field"],
  ["phyto_fields", "Phyto / certificate fields", "textarea"],
  ["kcb_fields", "KCB fields", "textarea"],
  ["certificate_fields", "Bio / origin certificate fields", "textarea"],
  ["value_tolerance", "Value tolerance"],
  ["weight_tolerance", "Weight tolerance kg"],
  ["quantity_tolerance", "Quantity tolerance"],
  ["packages_tolerance", "Packages tolerance"],
];

const UKDOCS_EXPECTED_COLUMNS = [
  "itemIdClientSystem", "itemNumber", "materialNumber", "invoiceIdClientSystem", "grossMassValue", "grossMassUnit",
  "netMassValue", "netMassUnit", "netPriceValue", "netPriceCurrencyIso", "value", "originCountryCode",
  "preferentialOriginCountryCode", "classificationType", "classificationValue", "goodsDescriptionText", "quantityValue",
  "quantityUnit", "packages", "order", "packageCode", "taricCode", "fullClassificationCode", "vbnCode", "vbnDescription",
];

function emptyUkdocsCustomer() {
  return {
    id: "",
    customer_name: "",
    match_hub_code: "",
    match_remark: "",
    customer_address: "",
    vat_number: "",
    eori_number: "",
    importer_number: "",
    default_owner: "",
    default_importer: "",
    default_delivery_terms: "",
    default_city: "",
    default_uk_arrival_port: "",
    default_currency: "",
    ready_email_subject: "",
    ready_email_body: "",
    default_invoice_language_text: "",
    default_document_references: "",
    show_invoice_vat_number: true,
    show_invoice_eori_number: true,
    show_invoice_importer_number: true,
    required_phyto: true,
    required_export_extra: false,
    required_generated_export: true,
    required_generated_invoices: true,
    menu_show_ukdocsinspection_inspection_list: true,
    menu_show_ukdocsinspection_locations_file: true,
    menu_show_ukdocsinspection_phyto: false,
    menu_show_ukdocsinspection_export_extra: false,
    menu_show_ukdocsinspection_generated_invoices: false,
    menu_show_ukdocsinspection_generated_export: false,
    menu_show_ukdocsprint_phyto: true,
    menu_show_ukdocsprint_export_extra: true,
    menu_show_ukdocsprint_generated_invoices: true,
    menu_show_ukdocsprint_generated_export: false,
    menu_show_ukdocsprint_inspection_list: false,
    menu_show_ukdocsprint_locations_file: false,
    export_defaults: Object.fromEntries(UKDOCS_EXPORT_DEFAULT_FIELDS.map(([key]) => [key, ""])),
  };
}

function emptyUkdocsShipmentDraft() {
  return {
    id: "",
    customer_id: "",
    shipment_date: new Date().toISOString().slice(0, 10),
    truck_number: "",
    trailer_number: "",
    invoice_numbers: "",
    invoice_numbers_by_category: Object.fromEntries(UKDOCS_CATEGORY_DEFINITIONS.map((category) => [category.code, ""])),
    export_reference: "",
    currency: "GBP",
    delivery_terms: "",
    uk_arrival_port: "",
    transport_customs_info: "",
    owner: "",
    regulation: "Export",
    destination_country: "GB / United Kingdom",
    customs_office_of_exit: "",
    location: "",
    delivery_terms_city: "",
    border_transport_mode: "Road",
    border_transport_nationality: "NL",
    importer: "",
    vessel: "",
    freight_costs: "",
    insurance: "",
    marks_and_numbers: "",
    container_number: "",
    uploaded_files: Object.fromEntries(UKDOCS_CATEGORY_DEFINITIONS.map((category) => [category.code, { category: category.code, file_name: "", uploaded_at: "", size: 0 }])),
    validation_warnings: [],
    audit_status: "",
    ready: false,
    notes: "",
    print_collection_id: "",
    reference_connect: "",
  };
}

function ukdocsStatusDefinition(status) {
  switch (status) {
    case "files_uploaded":
      return { label: "Files uploaded", tone: "info" };
    case "validated":
      return { label: "Ready to continue", tone: "info" };
    case "audit_passed":
      return { label: "Audit passed", tone: "success" };
    case "ready":
      return { label: "Ready", tone: "success" };
    case "failed":
      return { label: "Failed", tone: "danger" };
    default:
      return { label: "Not started", tone: "muted" };
  }
}

function ukdocsCombinedInvoiceNumbers(invoiceNumbersByCategory, uploadedFiles) {
  return UKDOCS_CATEGORY_DEFINITIONS
    .filter((category) => uploadedFiles?.[category.code]?.file_name || String(invoiceNumbersByCategory?.[category.code] || "").trim())
    .map((category) => String(invoiceNumbersByCategory?.[category.code] || "").trim())
    .filter(Boolean)
    .join("/");
}

function mergeUkdocsExportDefaults(baseDefaults = {}, overrideDefaults = {}) {
  const next = { ...baseDefaults };
  for (const [key] of UKDOCS_EXPORT_DEFAULT_FIELDS) {
    if (String(overrideDefaults?.[key] || "").trim()) {
      next[key] = overrideDefaults[key];
    }
  }
  return next;
}

function normalizeUkdocsMatchToken(value) {
  return String(value || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
}

function ukdocsMatchLines(value) {
  return String(value || "")
    .split(/\r?\n+/)
    .map(normalizeUkdocsMatchToken)
    .filter(Boolean);
}

function findUkdocsCustomerMatch(customers, collection) {
  const hubCode = normalizeUkdocsMatchToken(collection?.hub_code);
  const remark = normalizeUkdocsMatchToken(collection?.remark);
  let bestMatch = null;
  let bestScore = 0;
  for (const customer of customers || []) {
    if (!String(customer?.customer_name || "").trim()) {
      continue;
    }
    const customerHubCodes = ukdocsMatchLines(customer?.match_hub_code);
    const customerRemarks = ukdocsMatchLines(customer?.match_remark);
    if (!customerHubCodes.length && !customerRemarks.length) {
      continue;
    }
    let score = 0;
    if (customerHubCodes.length) {
      if (!hubCode || !customerHubCodes.includes(hubCode)) {
        continue;
      }
      score += 2;
    }
    if (customerRemarks.length) {
      if (!remark || !customerRemarks.some((item) => remark.includes(item))) {
        continue;
      }
      score += 1;
    }
    if (score > bestScore) {
      bestScore = score;
      bestMatch = customer;
    }
  }
  return bestMatch;
}

function ukdocsPrintSplitTokens(value) {
  return String(value || "")
    .split(/[\/,\s;]+/)
    .map((item) => String(item || "").trim())
    .filter(Boolean);
}

function ukdocsPrintInspectionMode(collection) {
  const pdType = String(collection?.pd_type || "").trim().toLowerCase();
  const pdTypeCompact = pdType.replace(/\s+/g, "");
  if (pdTypeCompact.includes("nakeuring")) {
    return "reinspection";
  }
  if (pdTypeCompact.includes("voorraadkeuring")) {
    return "";
  }
  if (pdTypeCompact === "voorraad") {
    return "stock_control";
  }
  if (!pdTypeCompact && collection?.collection_type === "stock_control") {
    return "stock_control";
  }
  return "";
}

function ukdocsPrintCollectionCustomer(collection, customers) {
  return (collection?.customer_id && (customers || []).find((item) => item.id === collection.customer_id))
    || findUkdocsCustomerMatch(customers || [], collection)
    || null;
}

function ukdocsInspectionDocumentKeys(collection) {
  const inspectionMode = ukdocsPrintInspectionMode(collection);
  if (inspectionMode === "stock_control") {
    return ["inspection_list", "locations_file"];
  }
  if (inspectionMode === "reinspection") {
    return ["phyto", "inspection_list", "export_extra"];
  }
  return ["phyto", "export_extra"];
}

function ukdocsMenuDocumentVisibility(customer, menuKey) {
  if (menuKey === "ukdocsinspection") {
    return {
      phyto: customer?.menu_show_ukdocsinspection_phyto === true,
      export_extra: customer?.menu_show_ukdocsinspection_export_extra === true,
      inspection_list: customer?.menu_show_ukdocsinspection_inspection_list !== false,
      locations_file: customer?.menu_show_ukdocsinspection_locations_file !== false,
      generated_invoice: customer?.menu_show_ukdocsinspection_generated_invoices === true,
      generated_export: customer?.menu_show_ukdocsinspection_generated_export === true,
    };
  }
  return {
    phyto: customer?.menu_show_ukdocsprint_phyto !== false,
    export_extra: customer?.menu_show_ukdocsprint_export_extra !== false,
    inspection_list: customer?.menu_show_ukdocsprint_inspection_list === true,
    locations_file: customer?.menu_show_ukdocsprint_locations_file === true,
    generated_invoice: customer?.menu_show_ukdocsprint_generated_invoices !== false,
    generated_export: customer?.menu_show_ukdocsprint_generated_export === true,
  };
}

function ukdocsPrintCollectionProgress(collection, customers) {
  const inspectionMode = ukdocsPrintInspectionMode(collection);
  if (inspectionMode === "stock_control") {
    const missing = [];
    if (!collection?.documents?.inspection_list?.storage_name) {
      missing.push("Inspection list");
    }
    if (!collection?.documents?.locations_file?.storage_name) {
      missing.push("Locations file");
    }
    const hasAny = Boolean(collection?.documents?.inspection_list?.storage_name || collection?.documents?.locations_file?.storage_name);
    return {
      customer: null,
      missing,
      status: missing.length === 0 ? "complete" : (hasAny ? "partial" : "pending"),
    };
  }
  if (inspectionMode === "reinspection") {
    const customer = ukdocsPrintCollectionCustomer(collection, customers);
    const missing = [];
    const phytoCount = (collection?.documents?.phyto_files || []).length;
    const phytoExpected = ukdocsPrintSplitTokens(collection?.reference_connect).length;
    const generatedFiles = collection?.documents?.generated_files || [];
    const generatedExportReady = generatedFiles.some((file) => file.document_kind === "export");
    const generatedInvoiceCount = generatedFiles.filter((file) => file.document_kind === "invoice").length;
    const invoiceExpected = ukdocsPrintSplitTokens(collection?.invoice_numbers).length;

    if (customer?.required_phyto !== false && phytoExpected > 0 && phytoCount < phytoExpected) {
      missing.push(`Phyto ${phytoCount}/${phytoExpected}`);
    }
    if (customer?.required_export_extra === true && !collection?.documents?.export_extra?.storage_name) {
      missing.push("Second export file");
    }
    if (customer?.required_generated_export !== false && !generatedExportReady) {
      missing.push("Generated export");
    }
    if (customer?.required_generated_invoices !== false) {
      if (invoiceExpected === 0) {
        missing.push("Invoice numbers");
      } else if (generatedInvoiceCount < invoiceExpected) {
        missing.push(`Invoices ${generatedInvoiceCount}/${invoiceExpected}`);
      }
    }
    if (!collection?.documents?.inspection_list?.storage_name) {
      missing.push("Inspection list");
    }

    return {
      customer,
      missing,
      status: missing.length === 0 ? "complete" : (phytoCount || generatedExportReady || generatedInvoiceCount || collection?.documents?.export_extra?.storage_name || collection?.documents?.inspection_list?.storage_name ? "partial" : "pending"),
    };
  }
  const customer = ukdocsPrintCollectionCustomer(collection, customers);
  const missing = [];
  const phytoCount = (collection?.documents?.phyto_files || []).length;
  const phytoExpected = ukdocsPrintSplitTokens(collection?.reference_connect).length;
  const generatedFiles = collection?.documents?.generated_files || [];
  const generatedExportReady = generatedFiles.some((file) => file.document_kind === "export");
  const generatedInvoiceCount = generatedFiles.filter((file) => file.document_kind === "invoice").length;
  const invoiceExpected = ukdocsPrintSplitTokens(collection?.invoice_numbers).length;

  if (customer?.required_phyto !== false && phytoExpected > 0 && phytoCount < phytoExpected) {
    missing.push(`Phyto ${phytoCount}/${phytoExpected}`);
  }
  if (customer?.required_export_extra === true && !collection?.documents?.export_extra?.storage_name) {
    missing.push("Second export file");
  }
  if (customer?.required_generated_export !== false && !generatedExportReady) {
    missing.push("Generated export");
  }
  if (customer?.required_generated_invoices !== false) {
    if (invoiceExpected === 0) {
      missing.push("Invoice numbers");
    } else if (generatedInvoiceCount < invoiceExpected) {
      missing.push(`Invoices ${generatedInvoiceCount}/${invoiceExpected}`);
    }
  }

  const complete = missing.length === 0 && (!!customer || !!collection?.customer_name || !!collection?.city_name);
  return {
    customer,
    missing,
    status: complete ? "complete" : (phytoCount || generatedExportReady || generatedInvoiceCount || collection?.documents?.export_extra?.storage_name ? "partial" : "pending"),
  };
}

function ukdocsGeneratedShipmentReady(collection) {
  const generatedFiles = collection?.documents?.generated_files || [];
  const generatedExportReady = generatedFiles.some((file) => file.document_kind === "export");
  const generatedInvoiceCount = generatedFiles.filter((file) => file.document_kind === "invoice").length;
  const invoiceExpected = ukdocsPrintSplitTokens(collection?.invoice_numbers).length;
  return generatedExportReady && invoiceExpected > 0 && generatedInvoiceCount >= invoiceExpected;
}

function ukdocsShipmentStatus(shipment) {
  const uploadedCount = Object.values(shipment?.uploaded_files || {}).filter((item) => item?.file_name).length;
  if (!uploadedCount) {
    return "not_started";
  }
  if (shipment?.audit_status === "failed") {
    return "failed";
  }
  if (shipment?.ready) {
    return "ready";
  }
  if (shipment?.audit_status === "passed") {
    return "audit_passed";
  }
  if (shipment?.customer_id && shipment?.shipment_date && shipment?.export_reference) {
    return "validated";
  }
  return "files_uploaded";
}

async function downloadUkdocsFilesWithPrompt(files) {
  if (!Array.isArray(files) || !files.length) {
    return;
  }
  if (typeof window.showDirectoryPicker !== "function") {
    files.forEach((file) => downloadBase64File(file.name, file.content_base64, file.mime_type));
    return;
  }

  try {
    const directoryHandle = await window.showDirectoryPicker();
    for (const file of files) {
      const fileHandle = await directoryHandle.getFileHandle(safeDownloadFilename(file.name), { create: true });
      const writable = await fileHandle.createWritable();
      const binary = window.atob(file.content_base64);
      const bytes = new Uint8Array(binary.length);
      for (let index = 0; index < binary.length; index += 1) {
        bytes[index] = binary.charCodeAt(index);
      }
      await writable.write(new Blob([bytes], { type: file.mime_type || "application/octet-stream" }));
      await writable.close();
    }
  } catch {
    files.forEach((file) => downloadBase64File(file.name, file.content_base64, file.mime_type));
  }
}

async function downloadUkdocsFileWithPrompt(file) {
  if (!file) {
    return;
  }
  if (typeof window.showSaveFilePicker !== "function") {
    downloadBase64File(file.name, file.content_base64, file.mime_type);
    return;
  }

  try {
    const handle = await window.showSaveFilePicker({
      suggestedName: safeDownloadFilename(file.name),
      types: [{
        description: "UKdocs file",
        accept: { [file.mime_type || "application/octet-stream"]: [".xlsx"] },
      }],
    });
    const writable = await handle.createWritable();
    const binary = window.atob(file.content_base64);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) {
      bytes[index] = binary.charCodeAt(index);
    }
    await writable.write(new Blob([bytes], { type: file.mime_type || "application/octet-stream" }));
    await writable.close();
  } catch {
    downloadBase64File(file.name, file.content_base64, file.mime_type);
  }
}

function UkdocsPage({ currentUser }) {
  const [activeMenu, setActiveMenu] = useState("new");
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [daySendingsBusy, setDaySendingsBusy] = useState(false);
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const [state, setState] = useState(null);
  const [customerDraft, setCustomerDraft] = useState(emptyUkdocsCustomer());
  const [shipmentDraft, setShipmentDraft] = useState(emptyUkdocsShipmentDraft());
  const [shipmentLoadDate, setShipmentLoadDate] = useState("");
  const [analysis, setAnalysis] = useState(null);
  const [generatedFiles, setGeneratedFiles] = useState([]);
  const [selectedAuditReportId, setSelectedAuditReportId] = useState("");
  const [exampleImportFiles, setExampleImportFiles] = useState({ invoice_example: null, export_example: null });
  const [shipmentUploadInputVersion, setShipmentUploadInputVersion] = useState(0);

  useEffect(() => {
    let cancelled = false;
    apiJson("/api/ukdocs/state")
      .then((payload) => {
        if (!cancelled) {
          setState(payload.state);
        }
      })
      .catch((loadError) => {
        if (!cancelled) {
          setError(loadError.message);
        }
      })
      .finally(() => {
        if (!cancelled) {
          setLoading(false);
        }
      });
    return () => {
      cancelled = true;
    };
  }, [currentUser]);

  const customers = state?.customers || [];
  const companySettings = state?.company_settings || {};
  const exportDefaults = state?.export_defaults || {};
  const templates = state?.templates || {};
  const columnMappings = state?.column_mappings || {};
  const shipments = state?.shipments || [];
  const auditReports = state?.audit_reports || [];
  const printCollections = state?.print_collections || [];
  const selectedAuditReport = auditReports.find((report) => report.id === selectedAuditReportId) || auditReports[0] || null;
  const selectedUkdocsCustomer = customers.find((item) => item.id === shipmentDraft.customer_id) || null;
  const selectedPrintCollection = printCollections.find((item) => item.id === shipmentDraft.print_collection_id) || null;
  const availablePrintCollections = useMemo(
    () => printCollections.filter((item) => String(item.shipment_date || "").slice(0, 10) === shipmentLoadDate && item.collection_type !== "stock_control"),
    [printCollections, shipmentLoadDate],
  );
  const hasSelectedShipmentContext = Boolean(shipmentDraft.id || shipmentDraft.print_collection_id);
  const activeExportDefaults = useMemo(
    () => mergeUkdocsExportDefaults(exportDefaults, selectedUkdocsCustomer?.export_defaults || {}),
    [exportDefaults, selectedUkdocsCustomer],
  );
  const combinedInvoiceNumbers = useMemo(
    () => ukdocsCombinedInvoiceNumbers(shipmentDraft.invoice_numbers_by_category, shipmentDraft.uploaded_files),
    [shipmentDraft.invoice_numbers_by_category, shipmentDraft.uploaded_files],
  );

  useEffect(() => {
    if (!state) {
      return;
    }
    setShipmentDraft((current) => ({
      ...current,
      currency: current.currency || activeExportDefaults.currency || "GBP",
      regulation: current.regulation || activeExportDefaults.regulation || "Export",
      destination_country: current.destination_country || activeExportDefaults.destination_country || "GB / United Kingdom",
      border_transport_mode: current.border_transport_mode || activeExportDefaults.border_transport_mode || "Road",
      border_transport_nationality: current.border_transport_nationality || activeExportDefaults.border_transport_nationality || "NL",
      customs_office_of_exit: current.customs_office_of_exit || activeExportDefaults.customs_office_of_exit || "",
      location: current.location || activeExportDefaults.location || "",
      delivery_terms_city: current.delivery_terms_city || activeExportDefaults.delivery_terms_city || "",
      freight_costs: current.freight_costs || activeExportDefaults.freight_costs || "",
      insurance: current.insurance || activeExportDefaults.insurance || "",
      importer: current.importer || selectedUkdocsCustomer?.default_importer || activeExportDefaults.importer_field || "",
      vessel: current.vessel || activeExportDefaults.vessel_field || "",
      owner: current.owner || selectedUkdocsCustomer?.default_owner || companySettings.company_name || "",
      invoice_numbers_by_category: current.invoice_numbers_by_category || Object.fromEntries(UKDOCS_CATEGORY_DEFINITIONS.map((category) => [category.code, ""])),
    }));
  }, [state, activeExportDefaults, companySettings]);

  function resetDrafts() {
    setShipmentDraft(emptyUkdocsShipmentDraft());
    setAnalysis(null);
    setGeneratedFiles([]);
    setShipmentUploadInputVersion((value) => value + 1);
  }

  function shipmentDraftHasWork(draft = shipmentDraft) {
    const uploadedAnyFile = Object.values(draft?.uploaded_files || {}).some((item) => item?.file_name);
    return Boolean(
      draft?.id
      || draft?.print_collection_id
      || draft?.customer_id
      || draft?.export_reference
      || draft?.reference_connect
      || draft?.truck_number
      || draft?.trailer_number
      || draft?.notes
      || uploadedAnyFile
      || analysis
      || generatedFiles.length,
    );
  }

  function buildDraftFromPrintCollection(collection, matchedCustomer = null) {
    const baseDraft = emptyUkdocsShipmentDraft();
    return {
      ...baseDraft,
      print_collection_id: collection?.id || "",
      customer_id: matchedCustomer?.id || "",
      shipment_date: collection?.shipment_date || baseDraft.shipment_date,
      truck_number: collection?.truck_number || "",
      trailer_number: collection?.trailer_number || "",
      reference_connect: collection?.reference_connect || "",
      delivery_terms: matchedCustomer?.default_delivery_terms || matchedCustomer?.export_defaults?.delivery_terms || baseDraft.delivery_terms,
      uk_arrival_port: matchedCustomer?.default_uk_arrival_port || baseDraft.uk_arrival_port,
      currency: matchedCustomer?.default_currency || matchedCustomer?.export_defaults?.currency || baseDraft.currency,
      owner: matchedCustomer?.default_owner || baseDraft.owner,
      importer: matchedCustomer?.default_importer || matchedCustomer?.importer_number || matchedCustomer?.eori_number || matchedCustomer?.export_defaults?.importer_field || baseDraft.importer,
      delivery_terms_city: matchedCustomer?.export_defaults?.delivery_terms_city || matchedCustomer?.default_city || baseDraft.delivery_terms_city,
      regulation: matchedCustomer?.export_defaults?.regulation || baseDraft.regulation,
      destination_country: matchedCustomer?.export_defaults?.destination_country || baseDraft.destination_country,
      customs_office_of_exit: matchedCustomer?.export_defaults?.customs_office_of_exit || baseDraft.customs_office_of_exit,
      location: matchedCustomer?.export_defaults?.location || baseDraft.location,
      border_transport_mode: matchedCustomer?.export_defaults?.border_transport_mode || baseDraft.border_transport_mode,
      border_transport_nationality: matchedCustomer?.export_defaults?.border_transport_nationality || baseDraft.border_transport_nationality,
      freight_costs: matchedCustomer?.export_defaults?.freight_costs || baseDraft.freight_costs,
      insurance: matchedCustomer?.export_defaults?.insurance || baseDraft.insurance,
      vessel: matchedCustomer?.export_defaults?.vessel_field || baseDraft.vessel,
      notes: [
        collection?.reference_connect ? `Reference connect: ${collection.reference_connect}` : "",
        collection?.city_name ? `Sending city: ${collection.city_name}` : "",
        collection?.hub_code ? `Hub: ${collection.hub_code}` : "",
        collection?.remark ? `Remark: ${collection.remark}` : "",
      ].filter(Boolean).join("\n"),
    };
  }

  function applyCustomerDefaults(customerId) {
    const customer = customers.find((item) => item.id === customerId);
    setShipmentDraft((current) => ({
      ...current,
      customer_id: customerId,
      delivery_terms: customer?.default_delivery_terms || customer?.export_defaults?.delivery_terms || current.delivery_terms,
      uk_arrival_port: customer?.default_uk_arrival_port || current.uk_arrival_port,
      currency: customer?.default_currency || customer?.export_defaults?.currency || current.currency,
      owner: customer?.default_owner || current.owner,
      importer: customer?.default_importer || customer?.importer_number || customer?.eori_number || customer?.export_defaults?.importer_field || current.importer,
      delivery_terms_city: customer?.export_defaults?.delivery_terms_city || customer?.default_city || current.delivery_terms_city,
      regulation: customer?.export_defaults?.regulation || current.regulation,
      destination_country: customer?.export_defaults?.destination_country || current.destination_country,
      customs_office_of_exit: customer?.export_defaults?.customs_office_of_exit || current.customs_office_of_exit,
      location: customer?.export_defaults?.location || current.location,
      border_transport_mode: customer?.export_defaults?.border_transport_mode || current.border_transport_mode,
      border_transport_nationality: customer?.export_defaults?.border_transport_nationality || current.border_transport_nationality,
      freight_costs: customer?.export_defaults?.freight_costs || current.freight_costs,
      insurance: customer?.export_defaults?.insurance || current.insurance,
      vessel: customer?.export_defaults?.vessel_field || current.vessel,
      notes: customer?.default_document_references || current.notes,
    }));
  }

  async function loadPrintCollectionsForDate(date, options = {}) {
    const normalizedDate = String(date || "").slice(0, 10);
    setShipmentLoadDate(normalizedDate);
    if (!normalizedDate) {
      if (options.resetDraft !== false) {
        resetDrafts();
      }
      return;
    }
    setDaySendingsBusy(true);
    setError("");
    if (!options.keepMessage) {
      setMessage("");
    }
    try {
      const payload = await apiJson("/api/ukdocs-print/sheet-sync", {
        method: "POST",
        body: JSON.stringify({ date: normalizedDate }),
      });
      setState((current) => ({
        ...current,
        print_collections: payload.print_collections || current?.print_collections || [],
      }));
      if (options.resetDraft !== false) {
        resetDrafts();
      }
      if (!options.silent) {
        setMessage(`Loaded ${payload.imported_count || 0} available sendings for ${payload.date}.`);
      }
    } catch (loadError) {
      setError(loadError.message);
    } finally {
      setDaySendingsBusy(false);
    }
  }

  function applyPrintCollection(collectionId) {
    const collection = printCollections.find((item) => item.id === collectionId) || null;
    if (!collection) {
      resetDrafts();
      return;
    }
    setShipmentLoadDate(String(collection.shipment_date || "").slice(0, 10));
    const matchedCustomer = (collection?.customer_id && customers.find((item) => item.id === collection.customer_id)) || findUkdocsCustomerMatch(customers, collection);
    const savedShipment = shipments.find((item) => item.print_collection_id === collectionId || item.id === collectionId) || null;
    const alreadyDone = Boolean(savedShipment || ukdocsGeneratedShipmentReady(collection));
    if (alreadyDone && !window.confirm("This sending already has saved UKdocs data. Opening or generating it again can replace the saved files and details. Do you want to continue?")) {
      return;
    }
    if (savedShipment) {
      selectShipment(savedShipment);
      return;
    }
    setShipmentDraft(buildDraftFromPrintCollection(collection, matchedCustomer));
    setAnalysis(null);
    setGeneratedFiles([]);
    setShipmentUploadInputVersion((value) => value + 1);
  }

  async function saveStatePatch(patch, successMessage) {
    setSaving(true);
    setError("");
    setMessage("");
    try {
      const payload = await apiJson("/api/ukdocs/state", {
        method: "PATCH",
        body: JSON.stringify(patch),
      });
      setState(payload.state);
      setMessage(successMessage);
    } catch (saveError) {
      setError(saveError.message);
    } finally {
      setSaving(false);
    }
  }

  function currentShipmentPayload() {
    return {
      ...shipmentDraft,
      invoice_numbers: combinedInvoiceNumbers,
      customers,
      company_settings: companySettings,
      export_defaults: activeExportDefaults,
      column_mappings: columnMappings,
      templates: state?.templates || {},
    };
  }

  async function analyzeShipment() {
    setSaving(true);
    setError("");
    setMessage("");
    try {
      const payload = await apiJson("/api/ukdocs/analyze", {
        method: "POST",
        body: JSON.stringify(currentShipmentPayload()),
      });
      setAnalysis(payload);
      setGeneratedFiles([]);
      setMessage(`UKdocs audit ${payload.audit.final_status}.`);
    } catch (analysisError) {
      setError(analysisError.message);
    } finally {
      setSaving(false);
    }
  }

  async function continueShipment() {
    await analyzeShipment();
  }

  async function generateDocuments() {
    setSaving(true);
    setError("");
    setMessage("");
    try {
      const payload = await apiJson("/api/ukdocs/generate", {
        method: "POST",
        body: JSON.stringify(currentShipmentPayload()),
      });
      setAnalysis(payload.analysis);
      setGeneratedFiles(payload.files || []);
      setState((current) => ({
        ...current,
        shipments: payload.shipments || current.shipments,
        audit_reports: payload.audit_reports || current.audit_reports,
        print_collections: payload.print_collections || current.print_collections,
      }));
      if (payload.audit_reports?.[0]?.id) {
        setSelectedAuditReportId(payload.audit_reports[0].id);
      }
      setShipmentDraft(payload.shipment || shipmentDraft);
      setMessage(`Generated ${payload.files?.length || 0} UKdocs files and saved the finished shipment.`);
      await downloadUkdocsFilesWithPrompt(payload.files || []);
    } catch (generateError) {
      setError(generateError.message);
    } finally {
      setSaving(false);
    }
  }

  async function deleteShipment(shipmentId) {
    if (!window.confirm("Delete this shipment draft from UKdocs history?")) {
      return;
    }
    setSaving(true);
    setError("");
    setMessage("");
    try {
      const payload = await apiJson(`/api/ukdocs/shipments/${encodeURIComponent(shipmentId)}`, { method: "DELETE" });
      setState((current) => ({ ...current, shipments: payload.shipments, print_collections: payload.print_collections || current.print_collections }));
      if (shipmentDraft.id === shipmentId) {
        resetDrafts();
      }
      setMessage("Shipment draft deleted.");
    } catch (deleteError) {
      setError(deleteError.message);
    } finally {
      setSaving(false);
    }
  }

  async function importCustomerFromExamples() {
    setSaving(true);
    setError("");
    setMessage("");
    try {
      const payload = await apiJson("/api/ukdocs/import-examples", {
        method: "POST",
        body: JSON.stringify({
          invoice_example: exampleImportFiles.invoice_example,
          export_example: exampleImportFiles.export_example,
        }),
      });
      const importedCustomer = {
        ...emptyUkdocsCustomer(),
        ...(payload.customer || {}),
        export_defaults: mergeUkdocsExportDefaults(emptyUkdocsCustomer().export_defaults, payload.export_defaults || {}),
      };
      const existingCustomer = customers.find((item) => item.customer_name && item.customer_name === importedCustomer.customer_name);
      const nextCustomer = {
        ...importedCustomer,
        id: existingCustomer?.id || `ukdocs-customer-${Date.now()}`,
      };
      const nextCustomers = existingCustomer
        ? customers.map((item) => (item.id === existingCustomer.id ? nextCustomer : item))
        : [nextCustomer, ...customers];
      const response = await apiJson("/api/ukdocs/state", {
        method: "PATCH",
        body: JSON.stringify({
          customers: nextCustomers,
          company_settings: payload.company_settings || {},
        }),
      });
      setState(response.state);
      setCustomerDraft(nextCustomer);
      setExampleImportFiles({ invoice_example: null, export_example: null });
      const warningText = (payload.warnings || []).filter(Boolean).join(" ");
      const importedName = nextCustomer.customer_name || "customer";
      setMessage(`Imported example data for ${importedName}.` + (warningText ? ` ${warningText}` : ""));
    } catch (importError) {
      setError(importError.message);
    } finally {
      setSaving(false);
    }
  }

  async function handleExampleImportFileChange(kind, file) {
    setExampleImportFiles((current) => ({
      ...current,
      [kind]: file
        ? {
            file_name: file.name,
            content_base64: "",
          }
        : null,
    }));
    if (!file) {
      return;
    }
    const contentBase64 = await fileToBase64(file);
    setExampleImportFiles((current) => ({
      ...current,
      [kind]: {
        file_name: file.name,
        content_base64: contentBase64,
      },
    }));
  }

  function updateCustomerExportDefault(key, value) {
    setCustomerDraft((current) => ({
      ...current,
      export_defaults: {
        ...(current.export_defaults || {}),
        [key]: value,
      },
    }));
  }

  function saveCustomer() {
    const nextCustomer = { ...customerDraft, id: customerDraft.id || `ukdocs-customer-${Date.now()}` };
    const nextCustomers = customerDraft.id
      ? customers.map((item) => (item.id === customerDraft.id ? nextCustomer : item))
      : [nextCustomer, ...customers];
    saveStatePatch({ customers: nextCustomers }, customerDraft.id ? "Customer updated." : "Customer added.");
    setCustomerDraft(emptyUkdocsCustomer());
  }

  function startEditCustomer(customer) {
    setCustomerDraft({ ...customer });
    setActiveMenu("customers");
  }

  function selectShipment(shipment) {
    setShipmentLoadDate(String(shipment?.shipment_date || "").slice(0, 10));
    setShipmentDraft({
      ...emptyUkdocsShipmentDraft(),
      ...shipment,
      invoice_numbers: shipment.invoice_numbers || ukdocsCombinedInvoiceNumbers(shipment.invoice_numbers_by_category, shipment.uploaded_files),
    });
    setAnalysis(null);
    setGeneratedFiles([]);
    setShipmentUploadInputVersion((value) => value + 1);
    setActiveMenu("new");
  }

  async function updateUploadedFile(categoryCode, file) {
    let nextFile = { category: categoryCode, file_name: "", size: 0, uploaded_at: "", content_base64: "" };
    if (file) {
      nextFile = {
        category: categoryCode,
        file_name: file.name,
        size: file.size || 0,
        uploaded_at: new Date().toISOString(),
        content_base64: await fileToBase64(file),
      };
    }
    setShipmentDraft((current) => ({
      ...current,
      uploaded_files: {
        ...(current.uploaded_files || {}),
        [categoryCode]: nextFile,
      },
    }));
    setAnalysis(null);
    setGeneratedFiles([]);
  }

  const uploadedCount = Object.values(shipmentDraft.uploaded_files || {}).filter((item) => item?.file_name).length;
  const canContinue = uploadedCount > 0 && !saving;
  const canGenerate = analysis?.audit?.final_status === "PASS" && !saving;

  function mappingText(categoryCode, columnName) {
    return ((columnMappings[categoryCode]?.aliases || {})[columnName] || []).join("\n");
  }

  function setMappingText(categoryCode, columnName, value) {
    const aliases = value.split(/\r?\n/).map((item) => item.trim()).filter(Boolean);
    setState((current) => ({
      ...current,
      column_mappings: {
        ...current.column_mappings,
        [categoryCode]: {
          ...(current.column_mappings?.[categoryCode] || { aliases: {} }),
          aliases: {
            ...((current.column_mappings?.[categoryCode] || {}).aliases || {}),
            [columnName]: aliases,
          },
        },
      },
    }));
  }

  function openNewShipmentScreen() {
    if (activeMenu === "new") {
      return;
    }
    if (shipmentDraftHasWork() && !window.confirm("Start a new UKdocs shipment screen and clear the current shipment view?")) {
      return;
    }
    resetDrafts();
    setShipmentLoadDate("");
    setMessage("");
    setError("");
    setActiveMenu("new");
  }

  if (loading) {
    return <div className="notice">Loading UKdocs workspace...</div>;
  }

  return (
    <section className="overview-stack ukdocs-page">
      <div className="tab-strip">
        {[
          ["new", "New shipment"],
          ["customers", "Customers"],
          ["company", "Company settings"],
          ["defaults", "Export defaults"],
          ["mappings", "Column mappings"],
          ["templates", "Templates"],
          ["history", "Shipment history"],
          ["audits", "Audit reports"],
        ].map(([key, label]) => (
          <button key={key} type="button" className={activeMenu === key ? "active" : ""} onClick={() => (key === "new" ? openNewShipmentScreen() : setActiveMenu(key))}>{label}</button>
        ))}
      </div>

      {message && <div className="notice">{message}</div>}
      {error && <div className="notice danger">{error}</div>}
      <div className="notice">Corrected uploaded dump files remain the source of truth. UKdocs now parses the real dump files, builds the control pivot, runs an audit, and generates first-pass Excel output files directly in Shadow App.</div>

      {activeMenu === "new" && (
        <div className="data-table-card ukdocs-stack">
          <div className="section-header">
            <h2>New shipment</h2>
            <div className={`ukdocs-status-badge ${ukdocsStatusDefinition(ukdocsShipmentStatus(shipmentDraft)).tone}`}>{ukdocsStatusDefinition(ukdocsShipmentStatus(shipmentDraft)).label}</div>
          </div>
          <div className="form-grid">
            <label>
              <span>Shipment date</span>
              <input
                type="date"
                value={shipmentLoadDate}
                onChange={async (event) => {
                  const nextDate = event.target.value;
                  if (nextDate === shipmentLoadDate) {
                    return;
                  }
                  if (shipmentDraftHasWork() && !window.confirm("Changing the day will clear the current shipment screen first. Continue?")) {
                    return;
                  }
                  await loadPrintCollectionsForDate(nextDate);
                }}
              />
            </label>
            <label>
              <span>&nbsp;</span>
              <button type="button" onClick={() => loadPrintCollectionsForDate(shipmentLoadDate, { resetDraft: false, keepMessage: true })} disabled={!shipmentLoadDate || daySendingsBusy || saving}>
                {daySendingsBusy ? "Loading..." : "Reload day sendings"}
              </button>
            </label>
            {!!shipmentLoadDate && (
              <label className="wide">
                <span>Available sending</span>
                <select value={shipmentDraft.print_collection_id || ""} onChange={(event) => applyPrintCollection(event.target.value)} disabled={daySendingsBusy}>
                  <option value="">Choose a sending from UKdocs Print</option>
                  {availablePrintCollections.map((collection) => {
                    const donePrefix = ukdocsGeneratedShipmentReady(collection) ? "[Done] " : "";
                    const remarkSuffix = String(collection.remark || "").trim() ? ` | ${String(collection.remark || "").trim()}` : "";
                    return (
                      <option key={collection.id} value={collection.id}>
                        {`${donePrefix}${collection.shipment_date || "-"} | ${collection.reference_connect || "-"} | ${collection.city_name || collection.customer_name || "-"} | ${collection.hub_code || "-"}${remarkSuffix}`}
                      </option>
                    );
                  })}
                </select>
              </label>
            )}
          </div>
          {!shipmentLoadDate && <div className="notice">Choose a date first. After that, UKdocs will show only the sendings available for that day.</div>}
          {!!shipmentLoadDate && !daySendingsBusy && !availablePrintCollections.length && !hasSelectedShipmentContext && <div className="notice">No UKdocs Print sendings are available for {shipmentLoadDate}.</div>}
          {!!shipmentLoadDate && !hasSelectedShipmentContext && !!availablePrintCollections.length && <div className="notice">Choose one of the {availablePrintCollections.length} available sendings for {shipmentLoadDate} to load its shipment details.</div>}
          {hasSelectedShipmentContext && (
            <>
              <div className="form-grid">
                <label><span>Customer / export user</span><select value={shipmentDraft.customer_id} onChange={(event) => applyCustomerDefaults(event.target.value)}><option value="">Choose customer</option>{customers.map((customer) => <option key={customer.id} value={customer.id}>{customer.customer_name}</option>)}</select></label>
                <label><span>Date</span><input type="date" value={shipmentDraft.shipment_date} onChange={(event) => setShipmentDraft({ ...shipmentDraft, shipment_date: event.target.value })} /></label>
            <label><span>Truck number</span><input value={shipmentDraft.truck_number || ""} onChange={(event) => setShipmentDraft({ ...shipmentDraft, truck_number: event.target.value })} placeholder="For example: 1 BHM" /></label>
            <label><span>Truck / trailer licence plate</span><input value={shipmentDraft.trailer_number} onChange={(event) => setShipmentDraft({ ...shipmentDraft, trailer_number: event.target.value })} placeholder="One licence plate for the whole export" /></label>
            <label><span>Combined export invoice numbers</span><input value={combinedInvoiceNumbers} readOnly placeholder="Filled automatically from the category invoice inputs below" /></label>
            <label><span>Export reference</span><input value={shipmentDraft.export_reference} onChange={(event) => setShipmentDraft({ ...shipmentDraft, export_reference: event.target.value })} /></label>
            <label><span>Reference connect</span><input value={shipmentDraft.reference_connect || ""} onChange={(event) => setShipmentDraft({ ...shipmentDraft, reference_connect: event.target.value })} placeholder="For example: 19053" /></label>
            <label><span>Currency</span><input value={shipmentDraft.currency} onChange={(event) => setShipmentDraft({ ...shipmentDraft, currency: event.target.value })} /></label>
            <label><span>Delivery terms</span><input value={shipmentDraft.delivery_terms} onChange={(event) => setShipmentDraft({ ...shipmentDraft, delivery_terms: event.target.value })} /></label>
            <label><span>UK arrival port</span><input value={shipmentDraft.uk_arrival_port} onChange={(event) => setShipmentDraft({ ...shipmentDraft, uk_arrival_port: event.target.value })} /></label>
            <label><span>Owner (export header)</span><input value={shipmentDraft.owner} onChange={(event) => setShipmentDraft({ ...shipmentDraft, owner: event.target.value })} placeholder="Exact owner text for the export file" /></label>
            <label><span>Regulation</span><input value={shipmentDraft.regulation} onChange={(event) => setShipmentDraft({ ...shipmentDraft, regulation: event.target.value })} /></label>
            <label><span>Country of destination</span><input value={shipmentDraft.destination_country} onChange={(event) => setShipmentDraft({ ...shipmentDraft, destination_country: event.target.value })} /></label>
            <label><span>Customs office of exit</span><input value={shipmentDraft.customs_office_of_exit} onChange={(event) => setShipmentDraft({ ...shipmentDraft, customs_office_of_exit: event.target.value })} /></label>
            <label><span>Location</span><input value={shipmentDraft.location} onChange={(event) => setShipmentDraft({ ...shipmentDraft, location: event.target.value })} /></label>
            <label><span>Delivery terms city</span><input value={shipmentDraft.delivery_terms_city} onChange={(event) => setShipmentDraft({ ...shipmentDraft, delivery_terms_city: event.target.value })} /></label>
            <label><span>Border transport mode</span><input value={shipmentDraft.border_transport_mode} onChange={(event) => setShipmentDraft({ ...shipmentDraft, border_transport_mode: event.target.value })} /></label>
            <label><span>Border transport nationality</span><input value={shipmentDraft.border_transport_nationality} onChange={(event) => setShipmentDraft({ ...shipmentDraft, border_transport_nationality: event.target.value })} /></label>
            <label><span>Importer number / reference</span><input value={shipmentDraft.importer} onChange={(event) => setShipmentDraft({ ...shipmentDraft, importer: event.target.value })} placeholder="Shown next to Importer in the export file" /></label>
            <label><span>Vessel</span><input value={shipmentDraft.vessel} onChange={(event) => setShipmentDraft({ ...shipmentDraft, vessel: event.target.value })} /></label>
            <label><span>Freight costs</span><input value={shipmentDraft.freight_costs} onChange={(event) => setShipmentDraft({ ...shipmentDraft, freight_costs: event.target.value })} /></label>
            <label><span>Insurance</span><input value={shipmentDraft.insurance} onChange={(event) => setShipmentDraft({ ...shipmentDraft, insurance: event.target.value })} /></label>
            <label><span>Marks and numbers</span><input value={shipmentDraft.marks_and_numbers} onChange={(event) => setShipmentDraft({ ...shipmentDraft, marks_and_numbers: event.target.value })} /></label>
            <label><span>Container number</span><input value={shipmentDraft.container_number} onChange={(event) => setShipmentDraft({ ...shipmentDraft, container_number: event.target.value })} /></label>
            <label className="wide"><span>Transport / customs information</span><textarea rows={4} value={shipmentDraft.transport_customs_info} onChange={(event) => setShipmentDraft({ ...shipmentDraft, transport_customs_info: event.target.value })} /></label>
            <label className="wide"><span>Notes / references</span><textarea rows={4} value={shipmentDraft.notes} onChange={(event) => setShipmentDraft({ ...shipmentDraft, notes: event.target.value })} /></label>
              </div>
              {selectedPrintCollection && <div className="notice">Selected sending: {selectedPrintCollection.city_name || "-"} | connect {selectedPrintCollection.reference_connect || "-"} | hub {selectedPrintCollection.hub_code || "-"} | PD {selectedPrintCollection.pd_form || "-"}</div>}

              <div className="ukdocs-upload-grid">
            {UKDOCS_CATEGORY_DEFINITIONS.map((category) => {
              const fileInfo = shipmentDraft.uploaded_files?.[category.code] || {};
              return (
                <div key={category.code} className="ukdocs-upload-card">
                  <strong>{category.shortLabel}</strong>
                  <input key={`${shipmentUploadInputVersion}-${category.code}`} type="file" accept=".xlsx,.xls,.csv" onChange={async (event) => updateUploadedFile(category.code, event.target.files?.[0] || null)} />
                  <label>
                    <span>Invoice number for this upload</span>
                    <input value={shipmentDraft.invoice_numbers_by_category?.[category.code] || ""} onChange={(event) => setShipmentDraft({ ...shipmentDraft, invoice_numbers_by_category: { ...(shipmentDraft.invoice_numbers_by_category || {}), [category.code]: event.target.value } })} placeholder={`External invoice number for ${category.label}`} />
                  </label>
                  <small>{fileInfo.file_name ? `${fileInfo.file_name} (${fileInfo.size || 0} bytes)` : "Optional category file not selected."}</small>
                </div>
              );
            })}
              </div>

              <div className="section-header"><h3>Workflow status</h3></div>
              <div className="ukdocs-badge-row">{["not_started", "files_uploaded", "audit_passed", "ready", "failed"].map((status) => { const definition = ukdocsStatusDefinition(status); return <div key={status} className={`ukdocs-status-badge ${definition.tone}`}>{definition.label}</div>; })}</div>
              <div className="row-actions spread-actions">
            <button type="button" onClick={resetDrafts} disabled={saving || daySendingsBusy}>Choose another sending</button>
            {!analysis && (
              <button type="button" className="primary" onClick={continueShipment} disabled={!canContinue}>
                {saving ? "Continuing..." : "Continue"}
              </button>
            )}
            {analysis && (
              <button type="button" className="primary" onClick={generateDocuments} disabled={!canGenerate}>
                {saving ? "Generating..." : "Generate documents"}
              </button>
            )}
            <button type="button" disabled={!generatedFiles.length} onClick={() => downloadUkdocsFilesWithPrompt(generatedFiles)}>Download files</button>
            <button type="button" disabled={!generatedFiles.length} onClick={() => generatedFiles.forEach((file) => downloadBase64File(file.name, file.content_base64, file.mime_type))}>Download only</button>
              </div>

              {analysis && (
                <div className="ukdocs-stack">
              <div className="section-header"><h3>Audit result</h3><div className={`ukdocs-status-badge ${analysis.audit.final_status === "PASS" ? "success" : "danger"}`}>{analysis.audit.final_status}</div></div>
              {!!analysis.audit.warnings?.length && <div className="notice danger">{analysis.audit.warnings.map((warning) => warning.message).join(" | ")}</div>}
              <div className="table-wrap">
                <table className="data-table">
                  <thead><tr><th>Category</th><th>Invoice</th><th>Rows</th><th>Quantity</th><th>Gross kg</th><th>Net kg</th><th>Packages</th><th>Value</th></tr></thead>
                  <tbody>
                    {analysis.categories.map((category) => <tr key={category.code}><td>{category.code} {category.label}</td><td>{category.invoice_number || "-"}</td><td>{category.row_count}</td><td>{category.totals.quantity}</td><td>{category.totals.gross_kg}</td><td>{category.totals.net_kg}</td><td>{category.totals.packages}</td><td>{category.totals.customs_value}</td></tr>)}
                    <tr className="summary-row"><td colSpan="3"><strong>Combined</strong></td><td><strong>{analysis.combined_totals.quantity}</strong></td><td><strong>{analysis.combined_totals.gross_kg}</strong></td><td><strong>{analysis.combined_totals.net_kg}</strong></td><td><strong>{analysis.combined_totals.packages}</strong></td><td><strong>{analysis.combined_totals.customs_value}</strong></td></tr>
                  </tbody>
                </table>
              </div>
              <div className="section-header"><h3>Control summary</h3></div>
              <div className="table-wrap">
                <table className="data-table">
                  <thead><tr><th>Scope</th><th>Group</th><th>Field</th><th>Dump</th><th>Invoice</th><th>Export</th><th>Invoice diff</th><th>Export diff</th><th>Status</th></tr></thead>
                  <tbody>
                    {(analysis.audit.summary_rows || []).map((row, index) => <tr key={`${row.scope}-${row.group_label}-${row.field}-${index}`}><td>{row.scope}</td><td>{row.group_label}</td><td>{row.field}</td><td>{row.dump_value || "0"}</td><td>{row.invoice_value || "-"}</td><td>{row.export_value || "-"}</td><td>{row.invoice_difference || "-"}</td><td>{row.export_difference || "-"}</td><td><span className={`ukdocs-status-badge ${row.status === "MATCH" ? "success" : "danger"}`}>{row.status}</span></td></tr>)}
                  </tbody>
                </table>
              </div>
              <div className="section-header"><h3>Control pivot</h3></div>
              <div className="table-wrap">
                <table className="data-table">
                  <thead><tr><th>Description</th><th>Quantity</th><th>Gross kg</th><th>Net kg</th><th>Packages</th><th>Value</th></tr></thead>
                  <tbody>{analysis.control_pivot_rows.map((row) => <tr key={row.description}><td>{row.description}</td><td>{row.quantity}</td><td>{row.gross_kg}</td><td>{row.net_kg}</td><td>{row.packages}</td><td>{row.customs_value}</td></tr>)}</tbody>
                </table>
              </div>
              <div className="section-header"><h3>Export preview</h3></div>
              <div className="table-wrap">
                <table className="data-table">
                  <thead><tr><th>Description</th><th>Commodity</th><th>Origin</th><th>Quantity</th><th>Net kg</th><th>Gross kg</th><th>Packages</th><th>Value</th></tr></thead>
                  <tbody>{analysis.export_rows.map((row, index) => <tr key={`${row.description}-${row.commodity_code}-${row.origin}-${index}`}><td>{row.description}</td><td>{row.commodity_code}</td><td>{row.origin}</td><td>{row.quantity}</td><td>{row.net_kg}</td><td>{row.gross_kg}</td><td>{row.packages}</td><td>{row.customs_value}</td></tr>)}</tbody>
                </table>
              </div>
              {!!generatedFiles.length && <div className="row-actions spread-actions">{generatedFiles.map((file) => <button key={file.name} type="button" onClick={() => downloadUkdocsFileWithPrompt(file)}>{file.name}</button>)}</div>}
                </div>
              )}
            </>
          )}
        </div>
      )}

      {activeMenu === "customers" && (
        <div className="data-table-card ukdocs-stack">
          <div className="section-header"><h2>Customers</h2></div>
          <div className="notice">Upload one invoice example and one export example to create or update a customer and prefill company and export-default data from your real UKdocs files.</div>
          <div className="ukdocs-upload-grid">
            <div className="ukdocs-upload-card">
              <strong>Invoice example</strong>
              <input type="file" accept=".xlsx,.xls" onChange={(event) => handleExampleImportFileChange("invoice_example", event.target.files?.[0] || null)} />
            </div>
            <div className="ukdocs-upload-card">
              <strong>Export example</strong>
              <input type="file" accept=".xlsx,.xls" onChange={(event) => handleExampleImportFileChange("export_example", event.target.files?.[0] || null)} />
            </div>
          </div>
          <div className="row-actions spread-actions"><button type="button" onClick={importCustomerFromExamples} disabled={saving || (!exampleImportFiles.invoice_example && !exampleImportFiles.export_example)}>{saving ? "Importing..." : "Import from example files"}</button></div>
          <div className="form-grid">
            {UKDOCS_CUSTOMER_FIELDS.map(([key, label, kind]) => (
              <label key={key} className={kind === "textarea" ? "wide" : ""}>
                <span>{label}</span>
                {kind === "textarea" ? <textarea rows={3} value={customerDraft[key] || ""} onChange={(event) => setCustomerDraft({ ...customerDraft, [key]: event.target.value })} /> : <input value={customerDraft[key] || ""} onChange={(event) => setCustomerDraft({ ...customerDraft, [key]: event.target.value })} />}
              </label>
            ))}
          </div>
          <div className="notice">Ready email templates can use placeholders like `{"{customer_name}"}`, `{"{shipment_date}"}`, `{"{city}"}`, `{"{reference_connect}"}`, `{"{invoice_numbers}"}`, `{"{truck_number}"}`, `{"{trailer_number}"}`, `{"{border_crossing}"}`, `{"{pd_form}"}`, `{"{re_export}"}`, `{"{pd_type}"}`, `{"{pd_code}"}`, and `{"{notes}"}`.</div>
          <div className="section-header"><h3>Invoice line visibility</h3></div>
          <div className="form-grid">
            {UKDOCS_CUSTOMER_INVOICE_VISIBILITY_FIELDS.map(([key, label]) => (
              <label key={key} className="checkbox-label">
                <input
                  type="checkbox"
                  checked={customerDraft[key] !== false}
                  onChange={(event) => setCustomerDraft({ ...customerDraft, [key]: event.target.checked })}
                />
                <span>{label}</span>
              </label>
            ))}
          </div>
          <div className="section-header"><h3>Customer export defaults</h3></div>
          <div className="form-grid">
            {UKDOCS_EXPORT_DEFAULT_FIELDS.map(([key, label, kind]) => (
              <label key={`customer-export-${key}`} className={kind === "textarea" ? "wide" : ""}>
                <span>{label}</span>
                {kind === "textarea" ? (
                  <textarea rows={3} value={customerDraft.export_defaults?.[key] || ""} onChange={(event) => updateCustomerExportDefault(key, event.target.value)} />
                ) : (
                  <input value={customerDraft.export_defaults?.[key] || ""} onChange={(event) => updateCustomerExportDefault(key, event.target.value)} />
                )}
              </label>
            ))}
          </div>
          <div className="row-actions spread-actions"><button type="button" className="primary" onClick={saveCustomer} disabled={saving}>{customerDraft.id ? "Update customer" : "Add customer"}</button><button type="button" onClick={() => setCustomerDraft(emptyUkdocsCustomer())} disabled={saving}>Clear form</button></div>
          <div className="checkbox-grid">
            {UKDOCS_CUSTOMER_REQUIRED_DOCUMENT_FIELDS.map(([key, label]) => (
              <label key={key} className="checkbox-field">
                <input
                  type="checkbox"
                  checked={customerDraft[key] === true}
                  onChange={(event) => setCustomerDraft({ ...customerDraft, [key]: event.target.checked })}
                />
                <span>{label}</span>
              </label>
            ))}
          </div>
          <div className="section-header"><h3>Menu document visibility</h3></div>
          <div className="checkbox-grid">
            {UKDOCS_CUSTOMER_MENU_DOCUMENT_FIELDS.map(([key, label]) => (
              <label key={key} className="checkbox-field">
                <input
                  type="checkbox"
                  checked={customerDraft[key] === true}
                  onChange={(event) => setCustomerDraft({ ...customerDraft, [key]: event.target.checked })}
                />
                <span>{label}</span>
              </label>
            ))}
          </div>
          <div className="table-wrap"><table className="data-table"><thead><tr><th>Name</th><th>Hub match</th><th>Remark match</th><th>Delivery terms</th><th>Ready mail template</th><th>UK port</th><th>Currency</th><th>VAT</th><th>Actions</th></tr></thead><tbody>{customers.map((customer) => <tr key={customer.id}><td>{customer.customer_name}</td><td>{customer.match_hub_code || "-"}</td><td>{customer.match_remark || "-"}</td><td>{customer.default_delivery_terms || customer.export_defaults?.delivery_terms || "-"}</td><td>{customer.ready_email_subject || customer.ready_email_body ? "Custom" : "Default"}</td><td>{customer.default_uk_arrival_port || "-"}</td><td>{customer.default_currency || customer.export_defaults?.currency || "-"}</td><td>{customer.vat_number || "-"}</td><td className="row-actions"><button type="button" onClick={() => startEditCustomer(customer)}>Edit</button></td></tr>)}{!customers.length && <tr><td colSpan="9">No UKdocs customers saved yet.</td></tr>}</tbody></table></div>
        </div>
      )}

      {activeMenu === "company" && (
        <div className="data-table-card ukdocs-stack">
          <div className="section-header"><h2>Company settings</h2></div>
          <div className="form-grid">{UKDOCS_COMPANY_FIELDS.map(([key, label, kind]) => <label key={key} className={kind === "textarea" ? "wide" : ""}><span>{label}</span>{kind === "textarea" ? <textarea rows={4} value={companySettings[key] || ""} onChange={(event) => setState((current) => ({ ...current, company_settings: { ...current.company_settings, [key]: event.target.value } }))} /> : <input value={companySettings[key] || ""} onChange={(event) => setState((current) => ({ ...current, company_settings: { ...current.company_settings, [key]: event.target.value } }))} />}</label>)}</div>
          <div className="row-actions spread-actions"><button type="button" className="primary" onClick={() => saveStatePatch({ company_settings: companySettings }, "Company settings saved.")} disabled={saving}>{saving ? "Saving..." : "Save company settings"}</button></div>
        </div>
      )}

      {activeMenu === "defaults" && (
        <div className="data-table-card ukdocs-stack">
          <div className="section-header"><h2>Global export defaults</h2></div>
          <div className="form-grid">{UKDOCS_EXPORT_DEFAULT_FIELDS.map(([key, label, kind]) => <label key={key} className={kind === "textarea" ? "wide" : ""}><span>{label}</span>{kind === "textarea" ? <textarea rows={3} value={exportDefaults[key] || ""} onChange={(event) => setState((current) => ({ ...current, export_defaults: { ...current.export_defaults, [key]: event.target.value } }))} /> : <input value={exportDefaults[key] || ""} onChange={(event) => setState((current) => ({ ...current, export_defaults: { ...current.export_defaults, [key]: event.target.value } }))} />}</label>)}</div>
          <div className="notice">These are fallback defaults and tolerances. A customer can override them with its own export defaults.</div>
          <div className="row-actions spread-actions"><button type="button" className="primary" onClick={() => saveStatePatch({ export_defaults: exportDefaults }, "Global export defaults saved.")} disabled={saving}>{saving ? "Saving..." : "Save global defaults"}</button></div>
        </div>
      )}

      {activeMenu === "mappings" && (
        <div className="data-table-card ukdocs-stack">
          <div className="section-header"><h2>Column mappings</h2></div>
          <div className="notice">Map source column names once per category. Small header variations can be stored here before the importer is connected to new corrected dumps.</div>
          {UKDOCS_CATEGORY_DEFINITIONS.map((category) => <div key={category.code} className="ukdocs-mapping-block"><div className="section-header"><h3>{category.code} {category.label}</h3></div><div className="table-wrap"><table className="data-table compact-table"><thead><tr><th>Expected column</th><th>Accepted aliases</th></tr></thead><tbody>{UKDOCS_EXPECTED_COLUMNS.map((columnName) => <tr key={`${category.code}-${columnName}`}><td>{columnName}</td><td><textarea rows={2} value={mappingText(category.code, columnName)} onChange={(event) => setMappingText(category.code, columnName, event.target.value)} placeholder="One alias per line" /></td></tr>)}</tbody></table></div></div>)}
          <div className="row-actions spread-actions"><button type="button" className="primary" onClick={() => saveStatePatch({ column_mappings: state.column_mappings }, "Column mappings saved.")} disabled={saving}>{saving ? "Saving..." : "Save column mappings"}</button></div>
        </div>
      )}

      {activeMenu === "templates" && (
        <div className="data-table-card ukdocs-stack">
          <div className="section-header"><h2>Templates</h2></div>
          <div className="form-grid"><label><span>Invoice template name</span><input value={templates.invoice_template_name || ""} onChange={(event) => setState((current) => ({ ...current, templates: { ...current.templates, invoice_template_name: event.target.value } }))} /></label><label><span>Export template name</span><input value={templates.export_template_name || ""} onChange={(event) => setState((current) => ({ ...current, templates: { ...current.templates, export_template_name: event.target.value } }))} /></label><label><span>Logo file name</span><input value={templates.logo_name || ""} onChange={(event) => setState((current) => ({ ...current, templates: { ...current.templates, logo_name: event.target.value } }))} /></label></div>
          <div className="notice">The first-pass generator already produces Excel files. This template section remains the place for future exact layout/template controls.</div>
          <div className="row-actions spread-actions"><button type="button" className="primary" onClick={() => saveStatePatch({ templates }, "Template references saved.")} disabled={saving}>{saving ? "Saving..." : "Save template references"}</button></div>
        </div>
      )}

      {activeMenu === "history" && (
        <div className="data-table-card ukdocs-stack">
          <div className="section-header"><h2>Shipment history</h2></div>
          <div className="notice">Only finished shipments are saved here automatically after successful document generation.</div>
          <div className="table-wrap"><table className="data-table"><thead><tr><th>Date</th><th>Reference</th><th>Customer</th><th>Invoices</th><th>Categories</th><th>Status</th><th>Updated</th><th>Actions</th></tr></thead><tbody>{shipments.map((shipment) => { const customer = customers.find((item) => item.id === shipment.customer_id); const status = ukdocsStatusDefinition(shipment.status || ukdocsShipmentStatus(shipment)); return <tr key={shipment.id}><td>{shipment.shipment_date || "-"}</td><td>{shipment.export_reference || "-"}</td><td>{customer?.customer_name || "-"}</td><td>{shipment.invoice_numbers || "-"}</td><td>{(shipment.categories_included || []).join(", ") || "-"}</td><td><span className={`ukdocs-status-badge ${status.tone}`}>{status.label}</span></td><td>{formatTimestamp(shipment.updated_at || shipment.created_at)}</td><td className="row-actions"><button type="button" onClick={() => selectShipment(shipment)}>Open</button><button type="button" onClick={() => deleteShipment(shipment.id)}>Delete</button></td></tr>; })}{!shipments.length && <tr><td colSpan="8">No finished UKdocs shipments saved yet.</td></tr>}</tbody></table></div>
        </div>
      )}

      {activeMenu === "audits" && (
        <div className="data-table-card ukdocs-stack">
          <div className="section-header"><h2>Audit reports</h2></div>
          <div className="table-wrap"><table className="data-table"><thead><tr><th>Shipment reference</th><th>Date</th><th>Customer</th><th>Created</th><th>Status</th><th>Warnings</th><th>Actions</th></tr></thead><tbody>{auditReports.map((report) => <tr key={report.id}><td>{report.shipment_reference || "-"}</td><td>{report.shipment_date || "-"}</td><td>{report.customer_name || "-"}</td><td>{formatTimestamp(report.created_at)}</td><td>{report.final_status || "Pending"}</td><td>{(report.warnings || []).join("; ") || "-"}</td><td className="row-actions"><button type="button" onClick={() => setSelectedAuditReportId(report.id)}>Open</button></td></tr>)}{!auditReports.length && <tr><td colSpan="7">No audit reports saved yet.</td></tr>}</tbody></table></div>
          {selectedAuditReport && (
            <>
              <div className="section-header"><h3>Audit detail</h3><div className={`ukdocs-status-badge ${selectedAuditReport.final_status === "PASS" ? "success" : "danger"}`}>{selectedAuditReport.final_status}</div></div>
              <p className="sidebar-note">{selectedAuditReport.summary || "-"}</p>
              <div className="table-wrap">
                <table className="data-table">
                  <thead><tr><th>Scope</th><th>Group</th><th>Field</th><th>Dump</th><th>Invoice</th><th>Export</th><th>Invoice diff</th><th>Export diff</th><th>Status</th></tr></thead>
                  <tbody>
                    {(selectedAuditReport.summary_rows || []).map((row, index) => <tr key={`${row.scope}-${row.group_label}-${row.field}-${index}`}><td>{row.scope}</td><td>{row.group_label}</td><td>{row.field}</td><td>{row.dump_value || "0"}</td><td>{row.invoice_value || "-"}</td><td>{row.export_value || "-"}</td><td>{row.invoice_difference || "-"}</td><td>{row.export_difference || "-"}</td><td><span className={`ukdocs-status-badge ${row.status === "MATCH" ? "success" : "danger"}`}>{row.status}</span></td></tr>)}
                  </tbody>
                </table>
              </div>
            </>
          )}
        </div>
      )}
    </section>
  );
}

const UKDOCS_PRINT_DOCUMENTS = [
  { key: "phyto", label: "Phytosanitaire document", accept: ".pdf,.xlsx,.xls" },
  { key: "export_extra", label: "Second export file", accept: ".pdf,.xlsx,.xls" },
  { key: "inspection_list", label: "Inspection list", accept: ".pdf,.xlsx,.xls" },
  { key: "locations_file", label: "Locations file", accept: ".pdf,.xlsx,.xls" },
];

function ukdocsPrintStatusDefinition(status) {
  switch (status) {
    case "complete":
      return { label: "Complete", tone: "success" };
    case "partial":
      return { label: "Partly collected", tone: "info" };
    default:
      return { label: "Waiting for files", tone: "muted" };
  }
}

function ukdocsCollectionDownloadEntries(collection, customer = null, menuKey = "all") {
  if (!collection?.id) {
    return [];
  }

  const collectionId = encodeURIComponent(collection.id);
  const entries = [];
  const seen = new Set();
  const phytoFiles = collection.documents?.phyto_files || [];
  const generatedFiles = collection.documents?.generated_files || [];
  const exportExtra = collection.documents?.export_extra || null;
  const inspectionList = collection.documents?.inspection_list || null;
  const locationsFile = collection.documents?.locations_file || null;
  const visibility = menuKey === "all"
    ? {
      phyto: true,
      export_extra: true,
      inspection_list: true,
      locations_file: true,
      generated_invoice: true,
      generated_export: true,
    }
    : ukdocsMenuDocumentVisibility(customer, menuKey);

  function pushEntry(prefix, file, href, fallbackLabel) {
    if (!file && !href) {
      return;
    }
    const normalizedName = String(file?.original_name || fallbackLabel || "").trim().toLowerCase();
    const identity = [
      prefix,
      normalizedName || String(file?.storage_name || "").trim().toLowerCase(),
    ].join("|");
    if (seen.has(identity)) {
      return;
    }
    seen.add(identity);
    entries.push({
      key: `${prefix}-${file?.storage_name || entries.length}`,
      label: file?.original_name || fallbackLabel,
      href,
    });
  }

  generatedFiles.forEach((generatedFile, index) => {
    const generatedKind = generatedFile.document_kind === "invoice" ? "generated_invoice" : (generatedFile.document_kind === "export" ? "generated_export" : "");
    if (generatedKind && visibility[generatedKind] !== true) {
      return;
    }
    pushEntry(
      "generated",
      generatedFile,
      `/api/ukdocs-print/collections/${collectionId}/documents/generated/${index}`,
      `Generated ${index + 1}`,
    );
  });

  if (visibility.phyto === true) {
  phytoFiles.forEach((phytoFile, index) => {
    pushEntry(
      "phyto",
      phytoFile,
      `/api/ukdocs-print/collections/${collectionId}/documents/phyto/${index}`,
      `Phyto ${index + 1}`,
    );
  });
  }

  if (visibility.export_extra === true && exportExtra?.storage_name) {
    pushEntry(
      "export-extra",
      exportExtra,
      `/api/ukdocs-print/collections/${collectionId}/documents/export_extra`,
      "Second export file",
    );
  }

  if (visibility.inspection_list === true && inspectionList?.storage_name) {
    pushEntry(
      "inspection-list",
      inspectionList,
      `/api/ukdocs-print/collections/${collectionId}/documents/inspection_list`,
      "Inspection list",
    );
  }

  if (visibility.locations_file === true && locationsFile?.storage_name) {
    pushEntry(
      "locations-file",
      locationsFile,
      `/api/ukdocs-print/collections/${collectionId}/documents/locations_file`,
      "Locations file",
    );
  }

  return entries;
}

async function downloadCollectionEntries(entries) {
  if (!Array.isArray(entries) || !entries.length) {
    return;
  }

  entries.forEach((entry) => {
    const link = document.createElement("a");
    link.href = entry.href;
    link.download = safeDownloadFilename(entry.label);
    document.body.appendChild(link);
    link.click();
    link.remove();
  });
}

function UkdocsPrintPage({ currentUser }) {
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [state, setState] = useState(null);
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const [selectedCollectionId, setSelectedCollectionId] = useState("");
  const [selectedCollectionDate, setSelectedCollectionDate] = useState(() => localDateIso());
  const [detailDrawerOpen, setDetailDrawerOpen] = useState(false);
  const [notesDraft, setNotesDraft] = useState("");
  const [gmailQuery, setGmailQuery] = useState("has:attachment newer_than:30d");
  const [gmailSyncResults, setGmailSyncResults] = useState([]);
  const [gmailBusy, setGmailBusy] = useState(false);
  const [gmailSettings, setGmailSettings] = useState({ gmail_connected_email: "" });
  const [sheetSyncDate, setSheetSyncDate] = useState(new Date().toISOString().slice(0, 10));
  const [sheetBusy, setSheetBusy] = useState(false);
  const [autoSyncEnabled, setAutoSyncEnabled] = useState(true);
  const canManageSettings = hasPermission(currentUser, PERMISSIONS.SETTINGS_MANAGE);

  useEffect(() => {
    let cancelled = false;
    Promise.all([
      apiJson("/api/ukdocs/state"),
      apiJson("/api/fust/settings").catch(() => ({ settings: { gmail_connected_email: "" } })),
    ])
      .then(([ukdocsPayload, settingsPayload]) => {
        if (!cancelled) {
          setState(ukdocsPayload.state);
          setGmailSettings(settingsPayload.settings || { gmail_connected_email: "" });
        }
      })
      .catch((loadError) => {
        if (!cancelled) {
          setError(loadError.message);
        }
      })
      .finally(() => {
        if (!cancelled) {
          setLoading(false);
        }
      });
    return () => {
      cancelled = true;
    };
  }, [currentUser]);

  const collections = state?.print_collections || [];
  const customers = state?.customers || [];
  const availableCollectionDates = useMemo(
    () => [...new Set(collections.map((item) => String(item.shipment_date || "").slice(0, 10)).filter(Boolean))].sort(),
    [collections],
  );
  const shipmentCollections = useMemo(
    () => collections.filter((item) => item.collection_type !== "stock_control"),
    [collections],
  );
  const filteredCollections = useMemo(
    () => shipmentCollections
      .filter((item) => String(item.shipment_date || "").slice(0, 10) === selectedCollectionDate)
      .sort((left, right) => {
        const leftRank = ukdocsPrintInspectionMode(left) ? 1 : 0;
        const rightRank = ukdocsPrintInspectionMode(right) ? 1 : 0;
        return leftRank - rightRank;
      }),
    [shipmentCollections, selectedCollectionDate],
  );
  const selectedCollection = filteredCollections.find((item) => item.id === selectedCollectionId || item.shipment_id === selectedCollectionId) || null;
  const selectedPhytoFiles = selectedCollection?.documents?.phyto_files || [];
  const selectedGeneratedFiles = selectedCollection?.documents?.generated_files || [];
  const selectedCollectionProgress = selectedCollection ? ukdocsPrintCollectionProgress(selectedCollection, customers) : null;

  useEffect(() => {
    setNotesDraft(selectedCollection?.notes || "");
  }, [selectedCollection?.id, selectedCollection?.notes]);

  useEffect(() => {
    if (!shipmentCollections.length) {
      setSelectedCollectionId("");
      setDetailDrawerOpen(false);
      return;
    }
    if (selectedCollectionId && !filteredCollections.some((item) => item.id === selectedCollectionId || item.shipment_id === selectedCollectionId)) {
      setSelectedCollectionId("");
      setDetailDrawerOpen(false);
    }
  }, [shipmentCollections, filteredCollections, selectedCollectionId]);

  function stepCollectionDate(days) {
    const current = new Date(`${selectedCollectionDate}T12:00:00`);
    current.setDate(current.getDate() + days);
    setSelectedCollectionDate(current.toISOString().slice(0, 10));
  }

  function openCollectionDetail(collectionId) {
    setSelectedCollectionId(collectionId);
    setDetailDrawerOpen(true);
  }

  function closeCollectionDetail() {
    setDetailDrawerOpen(false);
  }

  useEffect(() => {
    if (!autoSyncEnabled || !gmailSettings.gmail_connected_email) {
      return undefined;
    }
    const intervalId = window.setInterval(async () => {
      if (document.visibilityState !== "visible") {
        return;
      }
      try {
        const payload = await apiJson("/api/ukdocs-print/gmail/sync", {
          method: "POST",
          body: JSON.stringify({ query: gmailQuery, date: sheetSyncDate }),
        });
        setState((current) => ({ ...current, print_collections: payload.print_collections || current?.print_collections || [] }));
        setGmailSyncResults(payload.results || []);
      } catch {
      }
    }, 120000);
    return () => window.clearInterval(intervalId);
  }, [autoSyncEnabled, gmailSettings.gmail_connected_email, gmailQuery]);

  async function uploadCollectionFile(kind, file) {
    if (!selectedCollection || !file) {
      return;
    }
    setSaving(true);
    setMessage("");
    setError("");
    try {
      const contentBase64 = await fileToBase64(file);
      const payload = await apiJson(`/api/ukdocs-print/collections/${encodeURIComponent(selectedCollection.id)}/upload`, {
        method: "POST",
        body: JSON.stringify({
          kind,
          file: {
            file_name: file.name,
            mime_type: file.type || "application/octet-stream",
            content_base64: contentBase64,
          },
        }),
      });
      setState((current) => ({ ...current, print_collections: payload.print_collections || current?.print_collections || [] }));
      setMessage(kind === "phyto" ? "Phytosanitary document added." : `${UKDOCS_PRINT_DOCUMENTS.find((item) => item.key === kind)?.label || "File"} saved.`);
    } catch (uploadError) {
      setError(uploadError.message);
    } finally {
      setSaving(false);
    }
  }

  async function deleteCollectionDocument(kind, index = null) {
    if (!selectedCollection) {
      return;
    }
    if (!window.confirm("Delete this uploaded file?")) {
      return;
    }
    setSaving(true);
    setMessage("");
    setError("");
    try {
      const documentPath = index === null ? kind : `${kind}/${index}`;
      const payload = await apiJson(`/api/ukdocs-print/collections/${encodeURIComponent(selectedCollection.id)}/documents/${documentPath}`, {
        method: "DELETE",
      });
      setState((current) => ({ ...current, print_collections: payload.print_collections || current?.print_collections || [] }));
      setMessage("Uploaded file deleted.");
    } catch (deleteError) {
      setError(deleteError.message);
    } finally {
      setSaving(false);
    }
  }

  async function saveNotes() {
    if (!selectedCollection) {
      return;
    }
    setSaving(true);
    setMessage("");
    setError("");
    try {
      const payload = await apiJson(`/api/ukdocs-print/collections/${encodeURIComponent(selectedCollection.id)}`, {
        method: "PATCH",
        body: JSON.stringify({ notes: notesDraft }),
      });
      setState((current) => ({ ...current, print_collections: payload.print_collections || current?.print_collections || [] }));
      setMessage("Collection notes saved.");
    } catch (saveError) {
      setError(saveError.message);
    } finally {
      setSaving(false);
    }
  }

  async function sendReady(collectionId) {
    setSaving(true);
    setMessage("");
    setError("");
    try {
      const payload = await apiJson(`/api/ukdocs-print/collections/${encodeURIComponent(collectionId)}/send-ready`, {
        method: "POST",
      });
      setState((current) => ({ ...current, print_collections: payload.print_collections || current?.print_collections || [] }));
      setMessage(payload.delivery_email?.ok ? `Papers sent to ${payload.delivery_email.recipients.join(", ")}` : (payload.delivery_email?.error || "Could not send papers."));
    } catch (sendError) {
      setError(sendError.message);
    } finally {
      setSaving(false);
    }
  }

  async function deleteCollection(collectionId) {
    if (!window.confirm("Delete this UKdocs Print collection and its saved files?")) {
      return;
    }
    setSaving(true);
    setMessage("");
    setError("");
    try {
      const payload = await apiJson(`/api/ukdocs-print/collections/${encodeURIComponent(collectionId)}`, {
        method: "DELETE",
      });
      const nextCollections = payload.print_collections || [];
      const nextFilteredCollections = nextCollections
        .filter((item) => item.collection_type !== "stock_control")
        .filter((item) => String(item.shipment_date || "").slice(0, 10) === selectedCollectionDate);
      setState((current) => ({ ...current, print_collections: nextCollections }));
      setSelectedCollectionId(nextFilteredCollections[0]?.id || "");
      setNotesDraft(nextFilteredCollections[0]?.notes || "");
      if (!nextFilteredCollections.length || collectionId === selectedCollectionId) {
        setDetailDrawerOpen(false);
      }
      setMessage("Collection deleted.");
    } catch (deleteError) {
      setError(deleteError.message);
    } finally {
      setSaving(false);
    }
  }

  async function connectGmail() {
    setGmailBusy(true);
    setError("");
    setMessage("");
    try {
      const payload = await apiJson("/api/ukdocs-print/gmail/auth-url");
      window.location.href = payload.auth_url;
    } catch (connectError) {
      setError(connectError.message);
      setGmailBusy(false);
    }
  }

  async function syncSheetSendings() {
    setSheetBusy(true);
    setError("");
    setMessage("");
    try {
      const payload = await apiJson("/api/ukdocs-print/sheet-sync", {
        method: "POST",
        body: JSON.stringify({ date: sheetSyncDate }),
      });
      setState((current) => ({ ...current, print_collections: payload.print_collections || current?.print_collections || [] }));
      setMessage(`Loaded ${payload.imported_count || 0} sendings from ${payload.sheet_name || "spreadsheet"} for ${payload.date}.`);
    } catch (sheetError) {
      setError(sheetError.message);
    } finally {
      setSheetBusy(false);
    }
  }

  async function refreshReferenceConnectOnly() {
    setSheetBusy(true);
    setError("");
    setMessage("");
    try {
      const payload = await apiJson("/api/ukdocs-print/sheet-sync", {
        method: "POST",
        body: JSON.stringify({ date: sheetSyncDate, reference_connect_only: true }),
      });
      setState((current) => ({ ...current, print_collections: payload.print_collections || current?.print_collections || [] }));
      setMessage(`Reference connect refreshed for ${payload.updated_count || 0} shipments on ${payload.date}. Saved files and manual shipment data stayed untouched.`);
    } catch (sheetError) {
      setError(sheetError.message);
    } finally {
      setSheetBusy(false);
    }
  }

  async function syncGmail() {
    setGmailBusy(true);
    setError("");
    setMessage("");
    try {
      const payload = await apiJson("/api/ukdocs-print/gmail/sync", {
        method: "POST",
        body: JSON.stringify({ query: gmailQuery, date: sheetSyncDate }),
      });
      setState((current) => ({ ...current, print_collections: payload.print_collections || current?.print_collections || [] }));
      setGmailSyncResults(payload.results || []);
      setMessage(`Gmail sync finished for ${payload.date || sheetSyncDate}. ${payload.matched || 0} matched, ${payload.unmatched || 0} unmatched, ${payload.skipped || 0} skipped.`);
    } catch (syncError) {
      setError(syncError.message);
    } finally {
      setGmailBusy(false);
    }
  }

  if (loading) {
    return <div className="notice">Loading UKdocs Print workspace...</div>;
  }

  return (
    <section className="overview-stack ukdocs-page">
      {message && <div className="notice">{message}</div>}
      {error && <div className="notice danger">{error}</div>}
      <div className="notice">Use the spreadsheet to load the day sendings first. Then link UKdocs generated files to one of those sendings, and let Gmail attach the phytosanitary and second export documents onto the same record. Shipments with the same city, hub code, and remark are grouped together, so one shipment can hold multiple phyto references and files.</div>

      <div className="data-table-card ukdocs-stack">
        <div className="section-header"><h2>Today Sendings Spreadsheet</h2></div>
        <div className="form-grid">
          <label><span>Date to import</span><input type="date" value={sheetSyncDate} onChange={(event) => setSheetSyncDate(event.target.value)} /></label>
          <label><span>Spreadsheet settings</span><input value={gmailSettings.ukdocs_print_spreadsheet_id ? "Managed in Settings" : "Not set in Settings"} readOnly /></label>
        </div>
        <div className="row-actions spread-actions">
          <button type="button" className="primary" onClick={syncSheetSendings} disabled={sheetBusy}>{sheetBusy ? "Loading..." : "Load sendings from spreadsheet"}</button>
          <button type="button" onClick={refreshReferenceConnectOnly} disabled={sheetBusy}>{sheetBusy ? "Refreshing..." : "Refresh reference connect only"}</button>
        </div>
        <div className="notice">This imports the sending list for the selected day from the PD spreadsheet. Then UKdocs shipments can link to one of these sendings, and Gmail can match the phytosanitary PDF by reference connect.</div>
      </div>

      <div className="data-table-card ukdocs-stack">
        <div className="section-header"><h2>Gmail Inbox Pickup</h2></div>
        <div className="form-grid">
          <label><span>Connected Gmail account</span><input value={gmailSettings.gmail_connected_email || ""} readOnly placeholder="Not connected yet" /></label>
          <label><span>Export date</span><input type="date" value={sheetSyncDate} onChange={(event) => setSheetSyncDate(event.target.value)} /></label>
          <label className="wide"><span>Extra Gmail filter</span><input value={gmailQuery} onChange={(event) => setGmailQuery(event.target.value)} placeholder="has:attachment" /></label>
        </div>
        <div className="checkbox-grid">
          <label className="checkbox-field">
            <input type="checkbox" checked={autoSyncEnabled} onChange={(event) => setAutoSyncEnabled(event.target.checked)} />
            <span>Auto listen for Gmail attachments while this page is open</span>
          </label>
        </div>
        <div className="row-actions spread-actions">
          {canManageSettings && <button type="button" onClick={connectGmail} disabled={gmailBusy}>{gmailBusy ? "Connecting..." : "Connect Gmail"}</button>}
          <button type="button" className="primary" onClick={syncGmail} disabled={gmailBusy}>{gmailBusy ? "Syncing..." : "Sync Gmail attachments"}</button>
        </div>
      <div className="notice">The sync only checks emails for the selected export date, not the whole mailbox. It matches reference connect first, then invoice numbers, then truck or trailer registration. NVWA / e-CertNL emails are treated as phytosanitary documents automatically. Files only fill empty slots automatically, so manual uploads stay safe.</div>
        {!!gmailSyncResults.length && (
          <div className="table-wrap">
            <table className="data-table">
              <thead><tr><th>Status</th><th>File</th><th>Shipment</th><th>Type</th><th>Reason</th></tr></thead>
              <tbody>
                {gmailSyncResults.map((item, index) => <tr key={`${item.file_name}-${index}`}><td>{item.status}</td><td>{item.file_name || "-"}</td><td>{item.shipment_reference || "-"}</td><td>{item.kind || "-"}</td><td>{item.reason || "-"}</td></tr>)}
              </tbody>
            </table>
          </div>
        )}
      </div>

      <div className="ukdocs-print-layout">
        <div className="data-table-card ukdocs-stack">
          <div className="section-header"><h2>Collections</h2></div>
          <div className="form-grid">
            <label><span>Collection date</span><input type="date" value={selectedCollectionDate} onChange={(event) => setSelectedCollectionDate(event.target.value)} /></label>
            <label><span>Saved collection dates</span><input value={availableCollectionDates.length ? `${availableCollectionDates.length} day(s) saved` : "No saved dates yet"} readOnly /></label>
          </div>
          <div className="row-actions spread-actions">
            <button type="button" onClick={() => stepCollectionDate(-1)}>Previous day</button>
            <button type="button" className="primary" onClick={() => setSelectedCollectionDate(localDateIso())}>Today</button>
            <button type="button" onClick={() => stepCollectionDate(1)}>Next day</button>
          </div>
          <div className="notice">Showing collections for {selectedCollectionDate}. Saved shipments stay stored, but this view only shows the selected day.</div>
          <div className="ukdocs-upload-grid">
            {filteredCollections.map((collection) => {
              const progress = ukdocsPrintCollectionProgress(collection, customers);
              const status = ukdocsPrintStatusDefinition(progress.status);
              const isActive = detailDrawerOpen && selectedCollection?.id === collection.id;
              const downloadEntries = ukdocsCollectionDownloadEntries(collection, progress.customer, "ukdocsprint");
              const inspectionMode = ukdocsPrintInspectionMode(collection);
              const isStockControl = inspectionMode === "stock_control";
              return (
                <div key={collection.id} className={`ukdocs-upload-card ukdocs-collection-tile${isActive ? " active" : ""}`}>
                  <strong>{progress.customer?.customer_name || collection.customer_name || collection.city_name || "Shipment"}</strong>
                  <small>{collection.shipment_date || "-"}</small>
                  <small>{collection.city_name ? `City: ${collection.city_name}` : "City not linked yet"}</small>
                  <small>{collection.reference_connect ? `Connect: ${collection.reference_connect}` : "No connect ref yet"}</small>
                  <small>{collection.invoice_numbers ? `Invoices: ${collection.invoice_numbers}` : "No invoices linked yet"}</small>
                  <small>{collection.truck_number || collection.trailer_number ? `Truck: ${collection.truck_number || collection.trailer_number}` : "No truck linked yet"}</small>
                  <div className={`ukdocs-status-badge ${status.tone}`}>{progress.missing.length ? `${status.label} • ${progress.missing.join(", ")}` : status.label}</div>
                  {!!collection.delivery_email?.sent_at && <small>Sent {formatTimestamp(collection.delivery_email.sent_at)}</small>}
                  <div className="row-actions spread-actions">
                    <button type="button" className="primary" onClick={() => openCollectionDetail(collection.id)}>{isActive ? "Opened" : "Info"}</button>
                    {!isStockControl && !progress.missing.length && <button type="button" onClick={() => sendReady(collection.id)} disabled={saving}>Send papers</button>}
                    <button type="button" onClick={() => deleteCollection(collection.id)}>Delete</button>
                  </div>
                  {!!downloadEntries.length && (
                    <div className="ukdocs-download-box">
                      <div className="row-actions spread-actions">
                        <strong>Downloads</strong>
                        <button type="button" onClick={() => downloadCollectionEntries(downloadEntries)}>
                          Download all
                        </button>
                      </div>
                      <div className="ukdocs-download-list">
                        {downloadEntries.map((entry) => (
                          <a key={entry.key} href={entry.href} className="ukdocs-download-link">
                            {entry.label}
                          </a>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
              );
            })}
            {!shipmentCollections.length && <div className="notice">No spreadsheet sendings or UKdocs-linked collections are loaded yet.</div>}
            {!!shipmentCollections.length && !filteredCollections.length && <div className="notice">No collections saved for {selectedCollectionDate} yet.</div>}
          </div>
        </div>

        {detailDrawerOpen && <button type="button" className="ukdocs-drawer-backdrop" onClick={closeCollectionDetail} aria-label="Close shipment detail" />}
        <div className={`data-table-card ukdocs-stack ukdocs-drawer-panel${detailDrawerOpen ? " open" : ""}`}>
          <div className="section-header">
            <h2>Collection detail</h2>
            {selectedCollection && selectedCollectionProgress && <div className="row-actions"><div className={`ukdocs-status-badge ${ukdocsPrintStatusDefinition(selectedCollectionProgress.status).tone}`}>{ukdocsPrintStatusDefinition(selectedCollectionProgress.status).label}</div>{ukdocsPrintInspectionMode(selectedCollection) !== "stock_control" && !selectedCollectionProgress.missing.length && <button type="button" className="primary" onClick={() => sendReady(selectedCollection.id)} disabled={saving}>{saving ? "Sending..." : "Send papers ready"}</button>}<button type="button" onClick={closeCollectionDetail}>Close</button><button type="button" onClick={() => deleteCollection(selectedCollection.id)}>Delete</button></div>}
          </div>

          {!selectedCollection && <div className="notice">Tap Info on a shipment tile first.</div>}

          {selectedCollection && (
            <>
              <div className="form-grid">
                <label><span>Shipment reference</span><input value={selectedCollection.shipment_reference || ""} readOnly /></label>
                <label><span>Reference connect</span><input value={selectedCollection.reference_connect || ""} readOnly /></label>
                <label><span>Shipment date</span><input value={selectedCollection.shipment_date || ""} readOnly /></label>
                <label><span>Customer</span><input value={selectedCollectionProgress?.customer?.customer_name || selectedCollection.customer_name || ""} readOnly /></label>
                <label><span>City</span><input value={selectedCollection.city_name || ""} readOnly /></label>
                <label><span>Border crossing</span><input value={selectedCollection.border_crossing || ""} readOnly /></label>
                <label><span>Hub code</span><input value={selectedCollection.hub_code || ""} readOnly /></label>
                <label><span>Invoice numbers</span><input value={selectedCollection.invoice_numbers || ""} readOnly /></label>
                <label><span>Truck registration</span><input value={selectedCollection.truck_number || ""} readOnly /></label>
                <label><span>Trailer registration</span><input value={selectedCollection.trailer_number || ""} readOnly /></label>
                <label><span>PD form</span><input value={selectedCollection.pd_form || ""} readOnly /></label>
                <label><span>Re-Export</span><input value={selectedCollection.re_export || ""} readOnly /></label>
                <label><span>PD type</span><input value={selectedCollection.pd_type || ""} readOnly /></label>
                <label><span>PD code</span><input value={selectedCollection.pd_code || ""} readOnly /></label>
              </div>
              <div className="notice">{ukdocsPrintInspectionMode(selectedCollection) === "stock_control" ? (selectedCollectionProgress?.missing?.length ? `Still needed for stock control: ${selectedCollectionProgress.missing.join(", ")}` : "All stock control working papers are collected.") : (ukdocsPrintInspectionMode(selectedCollection) === "reinspection" ? (selectedCollectionProgress?.missing?.length ? `Still needed for nakeuring: ${selectedCollectionProgress.missing.join(", ")}. Gmail can match nakeuring files automatically when the reference, invoice, or truck/trailer matches.` : "All nakeuring inspection papers are collected. Gmail can match nakeuring files automatically when the reference, invoice, or truck/trailer matches.") : (selectedCollectionProgress?.missing?.length ? `Still needed: ${selectedCollectionProgress.missing.join(", ")}` : "All required files for this customer are collected."))}</div>
              {ukdocsPrintInspectionMode(selectedCollection) !== "stock_control" && !!selectedCollection?.delivery_email?.sent_at && <div className="notice">Last sent: {formatTimestamp(selectedCollection.delivery_email.sent_at)} to {(selectedCollection.delivery_email.recipients || []).join(", ") || "-"}</div>}
              {!!selectedCollection?.delivery_email?.error && !selectedCollection?.delivery_email?.ok && <div className="notice danger">{selectedCollection.delivery_email.error}</div>}

              <div className="ukdocs-upload-grid">
                {UKDOCS_PRINT_DOCUMENTS.filter((documentDefinition) => ukdocsInspectionDocumentKeys(selectedCollection).includes(documentDefinition.key)).map((documentDefinition) => {
                  const document = documentDefinition.key === "phyto"
                    ? null
                    : selectedCollection.documents?.[documentDefinition.key] || null;
                  return (
                    <div key={documentDefinition.key} className="ukdocs-upload-card">
                      <strong>{documentDefinition.label}</strong>
                      <input type="file" accept={documentDefinition.accept} onChange={(event) => uploadCollectionFile(documentDefinition.key, event.target.files?.[0] || null)} disabled={saving} />
                      {documentDefinition.key === "phyto" ? (
                        <>
                          <small>{selectedPhytoFiles.length ? `${selectedPhytoFiles.length} phytosanitary document(s) saved.` : "No file saved yet."}</small>
                          {!!selectedPhytoFiles.length && (
                            <div className="row-actions spread-actions">
                              {selectedPhytoFiles.map((phytoFile, index) => (
                                <span key={`${phytoFile.storage_name}-${index}`} className="row-actions spread-actions">
                                  <a href={`/api/ukdocs-print/collections/${encodeURIComponent(selectedCollection.id)}/documents/phyto/${index}`}>
                                    {phytoFile.original_name || `Phyto ${index + 1}`}
                                  </a>
                                  <button type="button" onClick={() => deleteCollectionDocument("phyto", index)} disabled={saving}>Delete</button>
                                </span>
                              ))}
                            </div>
                          )}
                        </>
                      ) : (
                        <>
                          <small>{document?.original_name ? `${document.original_name} saved ${formatTimestamp(document.saved_at)}` : "No file saved yet."}</small>
                          {document?.storage_name && <div className="row-actions"><a href={`/api/ukdocs-print/collections/${encodeURIComponent(selectedCollection.id)}/documents/${documentDefinition.key}`}>Download</a><button type="button" onClick={() => deleteCollectionDocument(documentDefinition.key)} disabled={saving}>Delete</button></div>}
                        </>
                      )}
                    </div>
                  );
                })}
              </div>

              {ukdocsPrintInspectionMode(selectedCollection) !== "stock_control" && <div className="ukdocs-upload-card">
                <strong>Generated UKdocs files</strong>
                <small>{selectedGeneratedFiles.length ? `${selectedGeneratedFiles.length} generated file(s) saved.` : "No generated files saved yet."}</small>
                {!!selectedGeneratedFiles.length && (
                  <div className="row-actions spread-actions">
                    {selectedGeneratedFiles.map((generatedFile, index) => (
                      <span key={`${generatedFile.storage_name}-${index}`} className="row-actions spread-actions">
                        <a href={`/api/ukdocs-print/collections/${encodeURIComponent(selectedCollection.id)}/documents/generated/${index}`}>
                          {generatedFile.original_name || `Generated file ${index + 1}`}
                        </a>
                        <button type="button" onClick={() => deleteCollectionDocument("generated", index)} disabled={saving}>Delete</button>
                      </span>
                    ))}
                  </div>
                )}
              </div>}

              <label className="wide">
                <span>Collection notes</span>
                <textarea rows={4} value={notesDraft} onChange={(event) => setNotesDraft(event.target.value)} placeholder="Optional notes about received emails or missing documents" />
              </label>
              <div className="row-actions spread-actions"><button type="button" className="primary" onClick={saveNotes} disabled={saving}>{saving ? "Saving..." : "Save notes"}</button></div>
            </>
          )}
        </div>
      </div>
    </section>
  );
}

function UkdocsInspectionPage({ currentUser }) {
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [state, setState] = useState(null);
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const [selectedCollectionId, setSelectedCollectionId] = useState("");
  const [detailDrawerOpen, setDetailDrawerOpen] = useState(false);
  const [notesDraft, setNotesDraft] = useState("");
  const selectedCollectionDate = localDateIso();

  useEffect(() => {
    let cancelled = false;
    apiJson("/api/ukdocs/state")
      .then((payload) => {
        if (!cancelled) {
          setState(payload.state);
        }
      })
      .catch((loadError) => {
        if (!cancelled) {
          setError(loadError.message);
        }
      })
      .finally(() => {
        if (!cancelled) {
          setLoading(false);
        }
      });
    return () => {
      cancelled = true;
    };
  }, [currentUser]);

  const collections = state?.print_collections || [];
  const customers = state?.customers || [];
  const inspectionCollections = useMemo(
    () => collections.filter((item) => Boolean(ukdocsPrintInspectionMode(item))),
    [collections],
  );
  const filteredCollections = useMemo(
    () => inspectionCollections
      .filter((item) => String(item.shipment_date || "").slice(0, 10) === selectedCollectionDate)
      .sort((left, right) => {
        const leftMode = ukdocsPrintInspectionMode(left);
        const rightMode = ukdocsPrintInspectionMode(right);
        const leftRank = leftMode === "stock_control" ? 0 : 1;
        const rightRank = rightMode === "stock_control" ? 0 : 1;
        return leftRank - rightRank;
      }),
    [inspectionCollections, selectedCollectionDate],
  );
  const selectedCollection = filteredCollections.find((item) => item.id === selectedCollectionId || item.shipment_id === selectedCollectionId) || null;
  const selectedCollectionProgress = selectedCollection ? ukdocsPrintCollectionProgress(selectedCollection, customers) : null;
  const selectedPhytoFiles = selectedCollection?.documents?.phyto_files || [];
  const selectedAllDownloadEntries = selectedCollection ? ukdocsCollectionDownloadEntries(selectedCollection, selectedCollectionProgress?.customer, "all") : [];

  useEffect(() => {
    setNotesDraft(selectedCollection?.notes || "");
  }, [selectedCollection?.id, selectedCollection?.notes]);

  useEffect(() => {
    if (!filteredCollections.length) {
      setSelectedCollectionId("");
      setDetailDrawerOpen(false);
      return;
    }
    if (selectedCollectionId && !filteredCollections.some((item) => item.id === selectedCollectionId || item.shipment_id === selectedCollectionId)) {
      setSelectedCollectionId("");
      setDetailDrawerOpen(false);
    }
  }, [filteredCollections, selectedCollectionId]);

  function openCollectionDetail(collectionId) {
    setSelectedCollectionId(collectionId);
    setDetailDrawerOpen(true);
  }

  function closeCollectionDetail() {
    setDetailDrawerOpen(false);
  }

  async function uploadCollectionFile(kind, file) {
    if (!selectedCollection || !file) {
      return;
    }
    setSaving(true);
    setMessage("");
    setError("");
    try {
      const contentBase64 = await fileToBase64(file);
      const payload = await apiJson(`/api/ukdocs-print/collections/${encodeURIComponent(selectedCollection.id)}/upload`, {
        method: "POST",
        body: JSON.stringify({
          kind,
          file: {
            file_name: file.name,
            mime_type: file.type || "application/octet-stream",
            content_base64: contentBase64,
          },
        }),
      });
      setState((current) => ({ ...current, print_collections: payload.print_collections || current?.print_collections || [] }));
      setMessage(`${UKDOCS_PRINT_DOCUMENTS.find((item) => item.key === kind)?.label || "File"} saved.`);
    } catch (uploadError) {
      setError(uploadError.message);
    } finally {
      setSaving(false);
    }
  }

  async function saveNotes() {
    if (!selectedCollection) {
      return;
    }
    setSaving(true);
    setMessage("");
    setError("");
    try {
      const payload = await apiJson(`/api/ukdocs-print/collections/${encodeURIComponent(selectedCollection.id)}`, {
        method: "PATCH",
        body: JSON.stringify({ notes: notesDraft }),
      });
      setState((current) => ({ ...current, print_collections: payload.print_collections || current?.print_collections || [] }));
      setMessage("Inspection notes saved.");
    } catch (saveError) {
      setError(saveError.message);
    } finally {
      setSaving(false);
    }
  }

  async function deleteCollection(collectionId) {
    if (!window.confirm("Delete this inspection shipment and its saved files?")) {
      return;
    }
    setSaving(true);
    setMessage("");
    setError("");
    try {
      const payload = await apiJson(`/api/ukdocs-print/collections/${encodeURIComponent(collectionId)}`, {
        method: "DELETE",
      });
      const nextCollections = payload.print_collections || [];
      const nextFilteredCollections = nextCollections
        .filter((item) => Boolean(ukdocsPrintInspectionMode(item)))
        .filter((item) => String(item.shipment_date || "").slice(0, 10) === selectedCollectionDate);
      setState((current) => ({ ...current, print_collections: nextCollections }));
      setSelectedCollectionId(nextFilteredCollections[0]?.id || "");
      setNotesDraft(nextFilteredCollections[0]?.notes || "");
      if (!nextFilteredCollections.length || collectionId === selectedCollectionId) {
        setDetailDrawerOpen(false);
      }
      setMessage("Inspection shipment deleted.");
    } catch (deleteError) {
      setError(deleteError.message);
    } finally {
      setSaving(false);
    }
  }

  if (loading) {
    return <div className="notice">Loading Phyto Inspection workspace...</div>;
  }

  return (
    <section className="overview-stack ukdocs-page">
      {message && <div className="notice">{message}</div>}
      {error && <div className="notice danger">{error}</div>}

      <div className="ukdocs-print-layout">
        <div className="data-table-card ukdocs-stack">
          <div className="section-header"><h2>Inspection shipments</h2></div>
          <div className="notice">Showing only today&apos;s inspection shipments from UKdocs Print collections.</div>
          <div className="ukdocs-collection-grid">
            {filteredCollections.map((collection) => {
              const progress = ukdocsPrintCollectionProgress(collection, customers);
              const status = ukdocsPrintStatusDefinition(progress.status);
              const isActive = detailDrawerOpen && selectedCollection?.id === collection.id;
              const downloadEntries = ukdocsCollectionDownloadEntries(collection, progress.customer, "ukdocsinspection");
              const inspectionMode = ukdocsPrintInspectionMode(collection);
              const title = inspectionMode === "stock_control" ? "Voorraad / stock control" : "Nakeuring";
              return (
                <div key={collection.id} className={`ukdocs-upload-card ukdocs-collection-tile${isActive ? " active" : ""}`}>
                  <strong>{title}</strong>
                  <small>{collection.shipment_date || "-"}</small>
                  <small>{collection.city_name ? `City: ${collection.city_name}` : "City not linked yet"}</small>
                  <small>{collection.reference_connect ? `Connect: ${collection.reference_connect}` : "No connect ref yet"}</small>
                  <small>{collection.invoice_numbers ? `Invoices: ${collection.invoice_numbers}` : "No invoices linked yet"}</small>
                  <small>{collection.truck_number || collection.trailer_number ? `Truck: ${collection.truck_number || collection.trailer_number}` : "No truck linked yet"}</small>
                  <div className={`ukdocs-status-badge ${status.tone}`}>{progress.missing.length ? `${status.label} • ${progress.missing.join(", ")}` : status.label}</div>
                  <div className="row-actions spread-actions">
                    <button type="button" className="primary" onClick={() => openCollectionDetail(collection.id)}>{isActive ? "Opened" : "Info"}</button>
                    <button type="button" onClick={() => deleteCollection(collection.id)}>Delete</button>
                  </div>
                  {!!downloadEntries.length && (
                    <div className="ukdocs-download-box">
                      <div className="row-actions spread-actions">
                        <strong>Downloads</strong>
                        <button type="button" onClick={() => downloadCollectionEntries(downloadEntries)}>Download all</button>
                      </div>
                      <div className="ukdocs-download-list">
                        {downloadEntries.map((entry) => (
                          <a key={entry.key} href={entry.href} className="ukdocs-download-link">
                            {entry.label}
                          </a>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
              );
            })}
            {!inspectionCollections.length && <div className="notice">No inspection-needed shipments are saved in UKdocs Print collections yet.</div>}
            {!!inspectionCollections.length && !filteredCollections.length && <div className="notice">No inspection shipments saved for {selectedCollectionDate} yet.</div>}
          </div>
        </div>

        {detailDrawerOpen && <button type="button" className="ukdocs-drawer-backdrop" onClick={closeCollectionDetail} aria-label="Close shipment detail" />}
        <div className={`data-table-card ukdocs-stack ukdocs-drawer-panel${detailDrawerOpen ? " open" : ""}`}>
          <div className="section-header">
            <h2>Inspection detail</h2>
            {selectedCollection && selectedCollectionProgress && <div className="row-actions"><div className={`ukdocs-status-badge ${ukdocsPrintStatusDefinition(selectedCollectionProgress.status).tone}`}>{ukdocsPrintStatusDefinition(selectedCollectionProgress.status).label}</div><button type="button" onClick={closeCollectionDetail}>Close</button><button type="button" onClick={() => deleteCollection(selectedCollection.id)}>Delete</button></div>}
          </div>

          {!selectedCollection && <div className="notice">Tap Info on an inspection shipment first.</div>}

          {selectedCollection && (
            <>
              <div className="form-grid">
                <label><span>Shipment reference</span><input value={selectedCollection.shipment_reference || ""} readOnly /></label>
                <label><span>Reference connect</span><input value={selectedCollection.reference_connect || ""} readOnly /></label>
                <label><span>Shipment date</span><input value={selectedCollection.shipment_date || ""} readOnly /></label>
                <label><span>City</span><input value={selectedCollection.city_name || ""} readOnly /></label>
                <label><span>Hub code</span><input value={selectedCollection.hub_code || ""} readOnly /></label>
                <label><span>Invoice numbers</span><input value={selectedCollection.invoice_numbers || ""} readOnly /></label>
                <label><span>Truck registration</span><input value={selectedCollection.truck_number || ""} readOnly /></label>
                <label><span>Trailer registration</span><input value={selectedCollection.trailer_number || ""} readOnly /></label>
                <label><span>PD form</span><input value={selectedCollection.pd_form || ""} readOnly /></label>
                <label><span>Re-Export</span><input value={selectedCollection.re_export || ""} readOnly /></label>
                <label><span>PD type</span><input value={selectedCollection.pd_type || ""} readOnly /></label>
                <label><span>PD code</span><input value={selectedCollection.pd_code || ""} readOnly /></label>
              </div>
              <div className="notice">{ukdocsPrintInspectionMode(selectedCollection) === "stock_control" ? (selectedCollectionProgress?.missing?.length ? `Still needed for voorraad / stock control: ${selectedCollectionProgress.missing.join(", ")}` : "All voorraad / stock control papers are collected.") : (selectedCollectionProgress?.missing?.length ? `Still needed for nakeuring: ${selectedCollectionProgress.missing.join(", ")}. Gmail can match nakeuring files automatically when the reference, invoice, or truck/trailer matches.` : "All nakeuring papers are collected. Gmail can match nakeuring files automatically when the reference, invoice, or truck/trailer matches.")}</div>

              <div className="ukdocs-upload-grid">
                {UKDOCS_PRINT_DOCUMENTS.filter((documentDefinition) => ukdocsInspectionDocumentKeys(selectedCollection).includes(documentDefinition.key)).map((documentDefinition) => {
                  const document = documentDefinition.key === "phyto"
                    ? null
                    : selectedCollection.documents?.[documentDefinition.key] || null;
                  return (
                    <div key={documentDefinition.key} className="ukdocs-upload-card">
                      <strong>{documentDefinition.label}</strong>
                      <input type="file" accept={documentDefinition.accept} onChange={(event) => uploadCollectionFile(documentDefinition.key, event.target.files?.[0] || null)} disabled={saving} />
                      {documentDefinition.key === "phyto" ? (
                        <>
                          <small>{selectedPhytoFiles.length ? `${selectedPhytoFiles.length} phytosanitary document(s) saved.` : "No file saved yet."}</small>
                          {!!selectedPhytoFiles.length && (
                            <div className="row-actions spread-actions">
                              {selectedPhytoFiles.map((phytoFile, index) => (
                                <span key={`${phytoFile.storage_name}-${index}`} className="row-actions spread-actions">
                                  <a href={`/api/ukdocs-print/collections/${encodeURIComponent(selectedCollection.id)}/documents/phyto/${index}`}>
                                    {phytoFile.original_name || `Phyto ${index + 1}`}
                                  </a>
                                  <button type="button" onClick={() => deleteCollectionDocument("phyto", index)} disabled={saving}>Delete</button>
                                </span>
                              ))}
                            </div>
                          )}
                        </>
                      ) : (
                        <>
                          <small>{document?.original_name ? `${document.original_name} saved ${formatTimestamp(document.saved_at)}` : "No file saved yet."}</small>
                          {document?.storage_name && <div className="row-actions"><a href={`/api/ukdocs-print/collections/${encodeURIComponent(selectedCollection.id)}/documents/${documentDefinition.key}`}>Download</a><button type="button" onClick={() => deleteCollectionDocument(documentDefinition.key)} disabled={saving}>Delete</button></div>}
                        </>
                      )}
                    </div>
                  );
                })}
              </div>

              {!!selectedAllDownloadEntries.length && (
                <div className="ukdocs-download-box">
                  <div className="row-actions spread-actions">
                    <strong>All linked files</strong>
                    <button type="button" onClick={() => downloadCollectionEntries(selectedAllDownloadEntries)}>Download all linked</button>
                  </div>
                  <div className="ukdocs-download-list">
                    {selectedAllDownloadEntries.map((entry) => (
                      <a key={entry.key} href={entry.href} className="ukdocs-download-link">
                        {entry.label}
                      </a>
                    ))}
                  </div>
                </div>
              )}

              <label className="wide">
                <span>Inspection notes</span>
                <textarea rows={5} value={notesDraft} onChange={(event) => setNotesDraft(event.target.value)} placeholder="Optional notes about inspection handling" />
              </label>
              <div className="row-actions spread-actions"><button type="button" className="primary" onClick={saveNotes} disabled={saving}>{saving ? "Saving..." : "Save notes"}</button></div>
            </>
          )}
        </div>
      </div>
    </section>
  );
}

function DagFoutjesPage() {
  return (
    <section className="overview-stack">
      <article className="panel" style={{ padding: 0, overflow: "hidden" }}>
        <iframe
          title="Dag Foutjes"
          src="/api/dag-foutjes/app"
          style={{
            display: "block",
            width: "100%",
            height: "calc(100vh - 10rem)",
            minHeight: "78vh",
            border: "none",
            background: "#fff",
          }}
        />
      </article>
    </section>
  );
}

function BunchesPage() {
  return (
    <section className="overview-stack">
      <article className="panel" style={{ padding: 0, overflow: "hidden" }}>
        <iframe
          title="Bunches"
          src="/api/bunches/app"
          style={{
            display: "block",
            width: "100%",
            height: "calc(100vh - 10rem)",
            minHeight: "78vh",
            border: "none",
            background: "#fff",
          }}
        />
      </article>
    </section>
  );
}

function foutenOverviewTypeSummary(entries) {
  const counts = new Map();
  for (const entry of entries) {
    const label = String(entry.type_label || entry.type_key || "-");
    counts.set(label, (counts.get(label) || 0) + 1);
  }
  return [...counts.entries()]
    .map(([label, count]) => ({ label, count }))
    .sort((left, right) => right.count - left.count || left.label.localeCompare(right.label));
}

function foutenOverviewPersonSummary(entries) {
  const people = new Map();
  for (const entry of entries) {
    const personKey = String(entry.person_id || entry.person_name || "");
    const current = people.get(personKey) || { person_name: entry.person_name || "-", total: 0, types: {} };
    current.total += 1;
    current.types[entry.type_label || entry.type_key || "-"] = (current.types[entry.type_label || entry.type_key || "-"] || 0) + 1;
    people.set(personKey, current);
  }
  return [...people.values()]
    .sort((left, right) => right.total - left.total || left.person_name.localeCompare(right.person_name));
}

function foutenOverviewPeriodSummary(entries, field) {
  const periods = new Map();
  for (const entry of entries) {
    const periodKey = String(entry[field] || "").trim();
    if (!periodKey) {
      continue;
    }
    const current = periods.get(periodKey) || { period: periodKey, total: 0, people: new Set(), types: {} };
    current.total += 1;
    current.people.add(entry.person_name || "-");
    current.types[entry.type_label || entry.type_key || "-"] = (current.types[entry.type_label || entry.type_key || "-"] || 0) + 1;
    periods.set(periodKey, current);
  }
  return [...periods.values()]
    .map((item) => ({
      period: item.period,
      total: item.total,
      people_count: item.people.size,
      top_type: Object.entries(item.types).sort((a, b) => b[1] - a[1])[0]?.[0] || "-",
    }))
    .sort((left, right) => right.period.localeCompare(left.period));
}

function FoutenOverviewPage() {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [fromDate, setFromDate] = useState("");
  const [toDate, setToDate] = useState("");
  const [personFilter, setPersonFilter] = useState("");

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError("");
    apiJson("/api/dag-foutjes/overview")
      .then((payload) => {
        if (cancelled) {
          return;
        }
        setData(payload);
      })
      .catch((nextError) => {
        if (!cancelled) {
          setError(nextError.message || String(nextError));
        }
      })
      .finally(() => {
        if (!cancelled) {
          setLoading(false);
        }
      });
    return () => {
      cancelled = true;
    };
  }, []);

  const entries = data?.entries || [];
  const filteredEntries = useMemo(() => {
    const personNeedle = personFilter.trim().toLowerCase();
    return entries.filter((entry) => {
      if (fromDate && String(entry.date || "") < fromDate) {
        return false;
      }
      if (toDate && String(entry.date || "") > toDate) {
        return false;
      }
      if (personNeedle && !String(entry.person_name || "").toLowerCase().includes(personNeedle)) {
        return false;
      }
      return true;
    });
  }, [entries, fromDate, toDate, personFilter]);

  const peopleSummary = useMemo(() => foutenOverviewPersonSummary(filteredEntries), [filteredEntries]);
  const typeSummary = useMemo(() => foutenOverviewTypeSummary(filteredEntries), [filteredEntries]);
  const daySummary = useMemo(() => foutenOverviewPeriodSummary(filteredEntries, "date"), [filteredEntries]);
  const weekSummary = useMemo(() => foutenOverviewPeriodSummary(filteredEntries, "iso_week"), [filteredEntries]);
  const monthSummary = useMemo(() => foutenOverviewPeriodSummary(filteredEntries, "month"), [filteredEntries]);

  return (
    <section className="overview-stack">
      <div className="data-table-card">
        <section className="toolbar" aria-label="Fouten overzicht filters">
          <label>
            <span>From date</span>
            <input type="date" value={fromDate} onChange={(event) => setFromDate(event.target.value)} />
          </label>
          <label>
            <span>To date</span>
            <input type="date" value={toDate} onChange={(event) => setToDate(event.target.value)} />
          </label>
          <label>
            <span>Search person</span>
            <input value={personFilter} onChange={(event) => setPersonFilter(event.target.value)} placeholder="Name" />
          </label>
          <Metric label="Mistakes" value={filteredEntries.length} />
          <Metric label="People" value={new Set(filteredEntries.map((entry) => entry.person_name)).size} />
          <Metric label="Types" value={typeSummary.length} />
        </section>
      </div>

      {loading && <div className="notice">Loading Fouten Overzicht...</div>}
      {error && <div className="notice danger">Unable to load Fouten Overzicht: {error}</div>}
      {!loading && !error && !filteredEntries.length && <div className="notice">No fouten records found for the selected filters.</div>}

      {!loading && !error && !!filteredEntries.length && (
        <>
          <div className="data-table-card">
            <h2>People ranking</h2>
            <div className="table-wrap">
              <table className="data-table">
                <thead>
                  <tr>
                    <th>Person</th>
                    <th>Total mistakes</th>
                    <th>Types</th>
                  </tr>
                </thead>
                <tbody>
                  {peopleSummary.map((row) => (
                    <tr key={row.person_name}>
                      <td>{row.person_name}</td>
                      <td>{row.total}</td>
                      <td>{Object.entries(row.types).sort((a, b) => b[1] - a[1]).map(([label, count]) => `${label}: ${count}`).join(", ")}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          <div className="data-table-card">
            <h2>Mistake types</h2>
            <div className="table-wrap">
              <table className="data-table">
                <thead>
                  <tr>
                    <th>Type</th>
                    <th>Total</th>
                  </tr>
                </thead>
                <tbody>
                  {typeSummary.map((row) => (
                    <tr key={row.label}>
                      <td>{row.label}</td>
                      <td>{row.count}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          <div className="data-table-card">
            <h2>Per day</h2>
            <div className="table-wrap">
              <table className="data-table">
                <thead>
                  <tr>
                    <th>Day</th>
                    <th>Total mistakes</th>
                    <th>People</th>
                    <th>Top type</th>
                  </tr>
                </thead>
                <tbody>
                  {daySummary.map((row) => (
                    <tr key={row.period}>
                      <td>{row.period}</td>
                      <td>{row.total}</td>
                      <td>{row.people_count}</td>
                      <td>{row.top_type}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          <div className="data-table-card">
            <h2>Per week</h2>
            <div className="table-wrap">
              <table className="data-table">
                <thead>
                  <tr>
                    <th>Week</th>
                    <th>Total mistakes</th>
                    <th>People</th>
                    <th>Top type</th>
                  </tr>
                </thead>
                <tbody>
                  {weekSummary.map((row) => (
                    <tr key={row.period}>
                      <td>{row.period}</td>
                      <td>{row.total}</td>
                      <td>{row.people_count}</td>
                      <td>{row.top_type}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          <div className="data-table-card">
            <h2>Per month</h2>
            <div className="table-wrap">
              <table className="data-table">
                <thead>
                  <tr>
                    <th>Month</th>
                    <th>Total mistakes</th>
                    <th>People</th>
                    <th>Top type</th>
                  </tr>
                </thead>
                <tbody>
                  {monthSummary.map((row) => (
                    <tr key={row.period}>
                      <td>{row.period}</td>
                      <td>{row.total}</td>
                      <td>{row.people_count}</td>
                      <td>{row.top_type}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </>
      )}
    </section>
  );
}

function App() {
  const [auth, setAuth] = useState({ loading: true, user: null, setupRequired: false });
  const rawPathname = typeof window !== "undefined" ? window.location.pathname : "/";
  const normalizedPathname = rawPathname !== "/" ? rawPathname.replace(/\/+$/, "").toLowerCase() : rawPathname;
  const publicClockMode = normalizedPathname === "/inklokken";
  const [page, setPage] = useState("dashboard");
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [fustMenuVersion, setFustMenuVersion] = useState(0);
  const loggedIn = Boolean(auth.user);
  const syncStatus = useSyncStatus(loggedIn);
  const [selectedDate, setSelectedDate] = useState("");
  const [dateWasManuallySelected, setDateWasManuallySelected] = useState(false);
  const [searchTerm, setSearchTerm] = useState("");
  const [expandedCustomers, setExpandedCustomers] = useState(() => new Set());
  const [lightbox, setLightbox] = useState(null);
  const [syncVersion, setSyncVersion] = useState(0);
  const requestedDate = dateWasManuallySelected ? selectedDate : "";
  const canViewPhotos = hasPermission(auth.user, PERMISSIONS.PHOTOS_VIEW);
  const visiblePages = availablePagesForUser(auth.user);
  const onPhotosPage = page === "dashboard" && canViewPhotos;
  const { data, loading, error } = useDashboardData(requestedDate, searchTerm, syncVersion, loggedIn && onPhotosPage);

  const dates = data?.dates || [];
  const activeDate = selectedDate || data?.selected_date || "";
  const firstDate = dates[0] || "";
  const lastDate = dates.at(-1) || "";
  const syncRunning = syncStatus?.state === "running";

  useEffect(() => {
    if (publicClockMode) {
      setAuth({ loading: false, user: null, setupRequired: false });
      return;
    }

    apiJson("/api/auth/me")
      .then((payload) => {
        const nextPage = defaultPageForUser(payload.user);
        setAuth({
          loading: false,
          user: payload.user,
          setupRequired: payload.setup_required,
        });
        if (payload.user) {
          setPage(nextPage);
          setSidebarOpen(true);
        }
      })
      .catch(() => setAuth({ loading: false, user: null, setupRequired: false }));
  }, [publicClockMode]);

  useEffect(() => {
    if (!dateWasManuallySelected && data?.selected_date) {
      setSelectedDate(data.selected_date);
    }
  }, [data?.selected_date, dateWasManuallySelected]);

  useEffect(() => {
    if (syncStatus?.state === "completed" || syncStatus?.state === "failed") {
      setSyncVersion((value) => value + 1);
    }
  }, [syncStatus?.state, syncStatus?.updated_at]);

  useEffect(() => {
    if (!auth.user) {
      return;
    }
    const fallbackPage = defaultPageForUser(auth.user);
    if (!visiblePages.some((item) => item.key === page)) {
      setPage(fallbackPage);
    }
  }, [auth.user, page, visiblePages]);

  const flatPhotoGroups = useMemo(() => {
    const groups = [];
    for (const group of data?.groups || []) {
      for (const run of group.runs || []) {
        if (Array.isArray(run.images) && run.images.length) {
          groups.push(run.images.map((image) => ({ image, run })));
        }
      }
    }
    return groups;
  }, [data]);

  async function startSync(pathname) {
    const params = new URLSearchParams();
    if (pathname === "/api/refresh-date") {
      params.set("date", activeDate);
    }

    const response = await fetch(`${pathname}?${params.toString()}`, { method: "POST" });
    if (response.ok) {
      setSyncVersion((value) => value + 1);
    }
  }

  async function handleLogout() {
    await apiJson("/api/auth/logout", { method: "POST" });
    setAuth({ loading: false, user: null, setupRequired: false });
    setPage("dashboard");
    setSidebarOpen(false);
  }

  function toggleCustomer(customerCode) {
    setExpandedCustomers((current) => {
      const next = new Set(current);
      if (next.has(customerCode)) {
        next.delete(customerCode);
      } else {
        next.add(customerCode);
      }
      return next;
    });
  }

  if (publicClockMode) {
    const heading = pageHeading("clock");
    return (
      <main className="workspace public-clock-workspace">
        <header className="page-header">
          <h1>{heading.title}</h1>
          {heading.caption ? <p>{heading.caption}</p> : null}
        </header>
        <ClockPage currentUser={null} publicMode />
      </main>
    );
  }

  if (auth.loading) {
    return <AuthShell title="Loading..." />;
  }

  if (!auth.user) {
    return (
      <AuthShell title="Sjaak vd Vijver App">
        {auth.setupRequired ? (
          <SetupForm onSetup={(user) => {
            setAuth({ loading: false, user, setupRequired: false });
            setPage(defaultPageForUser(user));
            setSidebarOpen(true);
          }}
          />
        ) : (
          <LoginForm onLogin={(user) => {
            setAuth({ loading: false, user, setupRequired: false });
            setPage(defaultPageForUser(user));
            setSidebarOpen(true);
          }}
          />
        )}
      </AuthShell>
    );
  }

  const heading = pageHeading(page);

  return (
    <>
      <button
        type="button"
        className="sidebar-toggle"
        aria-label={sidebarOpen ? "Close menu" : "Open menu"}
        aria-expanded={sidebarOpen}
        onClick={() => setSidebarOpen((open) => !open)}
      >
        <span />
        <span />
        <span />
      </button>
      {sidebarOpen && (
        <button
          type="button"
          className="sidebar-backdrop"
          aria-label="Close menu"
          onClick={() => setSidebarOpen(false)}
        />
      )}
      <aside className={`sidebar ${sidebarOpen ? "open" : ""}`}>
        <div>
          <h1>Sjaak vd Vijver</h1>
          <p className="eyebrow">Dashboard</p>
        </div>

        <nav className="side-nav" aria-label="Shadow app pages">
          {visiblePages.map((item) => (
            <button
              key={item.key}
              className={page === item.key ? "active" : ""}
              onClick={(event) => {
                event.currentTarget.blur();
                setPage(item.key);
                if (item.key === "fust") {
                  setFustMenuVersion((value) => value + 1);
                }
                setSidebarOpen(false);
              }}
            >
              {item.label}
            </button>
          ))}
        </nav>

        {canViewPhotos && (
          <>
            <div className="control-stack">
              <button className="primary" disabled={syncRunning} onClick={() => startSync("/api/rebuild")}>
                Rebuild run index
              </button>
              <button disabled={syncRunning || !activeDate} onClick={() => startSync("/api/refresh-date")}>
                Refresh this date
              </button>
              <p className="sidebar-note">
                {data?.generated_at
                  ? `Using saved run index from ${formatTimestamp(data.generated_at)}.`
                  : "No saved run index loaded yet."}
              </p>
            </div>

            <SyncPanel status={syncStatus} />
            <ParseErrors errors={data?.parse_errors || []} />
          </>
        )}

        <div className="signed-in">
          <p>Signed in as <strong>{auth.user.username}</strong></p>
          <button onClick={handleLogout}>Log out</button>
        </div>
      </aside>

      <main className="workspace">
        <header className="page-header">
          <h1>{heading.title}</h1>
          <p>{heading.caption}</p>
        </header>

        {page === "users" && <UsersPage currentUser={auth.user} />}
        {page === "fust" && <FustPage currentUser={auth.user} menuVersion={fustMenuVersion} />}
        {page === "cmrprint" && <CmrPrintPage currentUser={auth.user} />}
        {page === "clock" && <ClockPage currentUser={auth.user} />}
        {page === "hallocations" && <HalLocationsPage currentUser={auth.user} />}
        {page === "expeditionstickers" && <ExpeditionStickerPage currentUser={auth.user} />}
        {page === "bunches" && <BunchesPage currentUser={auth.user} />}
        {page === "dagfoutjes" && <DagFoutjesPage currentUser={auth.user} />}
        {page === "foutenoverzicht" && <FoutenOverviewPage currentUser={auth.user} />}
        {page === "ukdocsprint" && <UkdocsPrintPage currentUser={auth.user} />}
        {page === "ukdocsinspection" && <UkdocsInspectionPage currentUser={auth.user} />}
        {page === "settings" && <SettingsPage currentUser={auth.user} />}
        {page === "ukdocs" && <UkdocsPage currentUser={auth.user} />}
        {page === "dashboard" && canViewPhotos && (
          <>
            <section className="toolbar" aria-label="Filters">
              <label>
                <span>Filter by date</span>
                <input
                  type="date"
                  value={activeDate}
                  min={firstDate}
                  max={lastDate}
                  onChange={(event) => {
                    setDateWasManuallySelected(true);
                    setSelectedDate(event.target.value);
                  }}
                />
              </label>

              <label>
                <span>Search customer code</span>
                <input
                  value={searchTerm}
                  onChange={(event) => setSearchTerm(event.target.value)}
                  placeholder="cust123"
                />
              </label>

              <Metric label="Customers" value={data?.metrics?.customers || 0} />
              <Metric label="Runs" value={data?.metrics?.runs || 0} />
              <Metric label="Images" value={data?.metrics?.images || 0} />
            </section>

            {loading && <div className="notice">Loading run data...</div>}
            {error && <div className="notice danger">Unable to load data: {error}</div>}
            {data?.cache_missing && (
              <div className="notice">
                No saved run index found. Start a rebuild, or open the Streamlit app once to create the shared cache.
              </div>
            )}
            {data?.auto_sync && (
              <div className="notice">
                Automatic background sync started
                {data.auto_sync.date ? ` for ${data.auto_sync.date}` : ""}.
              </div>
            )}

            {!loading && !error && data && data.groups.length === 0 && (
              <div className="notice">No runs match the selected filters.</div>
            )}

            <section className="customer-list">
              {(data?.groups || []).map((group) => {
                const isExpanded = expandedCustomers.has(group.customer_code);
                return (
                  <CustomerGroup
                    key={group.customer_code}
                    group={group}
                    expanded={isExpanded}
                    onToggle={() => toggleCustomer(group.customer_code)}
                    onOpenPhoto={(run, imageIndex) => {
                      const photos = run.images.map((image) => ({ image, run }));
                      setLightbox({ photos, index: imageIndex });
                    }}
                  />
                );
              })}
            </section>
          </>
        )}
      </main>

      {lightbox && (
        <Lightbox
          photos={lightbox.photos}
          index={lightbox.index}
          onChange={(index) => setLightbox((current) => ({ ...current, index }))}
          onClose={() => setLightbox(null)}
        />
      )}

      <div className="sr-only" aria-live="polite">
        {flatPhotoGroups.length} photo groups loaded
      </div>
    </>
  );
}

function AuthShell({ title, children }) {
  return (
    <main className="auth-page">
      <section className="auth-panel">
        <p className="eyebrow">SnappySjaak</p>
        <h1>{title}</h1>
        {children}
      </section>
    </main>
  );
}

function LoginForm({ onLogin }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");

  async function submit(event) {
    event.preventDefault();
    setError("");
    try {
      const payload = await apiJson("/api/auth/login", {
        method: "POST",
        body: JSON.stringify({ username, password }),
      });
      onLogin(payload.user);
    } catch (loginError) {
      setError(loginError.message);
    }
  }

  return (
    <form className="auth-form" onSubmit={submit}>
      <label>
        <span>Username</span>
        <input value={username} onChange={(event) => setUsername(event.target.value)} autoComplete="username" />
      </label>
      <label>
        <span>Password</span>
        <input
          type="password"
          value={password}
          onChange={(event) => setPassword(event.target.value)}
          autoComplete="current-password"
        />
      </label>
      {error && <div className="notice danger">{error}</div>}
      <button className="primary" type="submit">Log in</button>
    </form>
  );
}

function SetupForm({ onSetup }) {
  const [username, setUsername] = useState("admin");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");

  async function submit(event) {
    event.preventDefault();
    setError("");
    try {
      const payload = await apiJson("/api/auth/setup", {
        method: "POST",
        body: JSON.stringify({ username, password }),
      });
      onSetup(payload.user);
    } catch (setupError) {
      setError(setupError.message);
    }
  }

  return (
    <form className="auth-form" onSubmit={submit}>
      <p className="sidebar-note">Create the first admin account for this shadow app.</p>
      <label>
        <span>Admin username</span>
        <input value={username} onChange={(event) => setUsername(event.target.value)} autoComplete="username" />
      </label>
      <label>
        <span>Admin password</span>
        <input
          type="password"
          value={password}
          onChange={(event) => setPassword(event.target.value)}
          autoComplete="new-password"
        />
      </label>
      {error && <div className="notice danger">{error}</div>}
      <button className="primary" type="submit">Create admin</button>
    </form>
  );
}

function UsersPage({ currentUser }) {
  const [users, setUsers] = useState([]);
  const [form, setForm] = useState({
    username: "",
    password: "",
    role: "viewer",
    permissions: [...DEFAULT_PERMISSIONS_BY_ROLE.viewer],
  });
  const [error, setError] = useState("");
  const [message, setMessage] = useState("");

  async function loadUsers() {
    const payload = await apiJson("/api/users");
    setUsers(payload.users || []);
  }

  useEffect(() => {
    loadUsers().catch((loadError) => setError(loadError.message));
  }, []);

  async function addUser(event) {
    event.preventDefault();
    setError("");
    setMessage("");
    try {
      await apiJson("/api/users", {
        method: "POST",
        body: JSON.stringify(form),
      });
      setForm({
        username: "",
        password: "",
        role: "viewer",
        permissions: [...DEFAULT_PERMISSIONS_BY_ROLE.viewer],
      });
      setMessage("User added.");
      await loadUsers();
    } catch (addError) {
      setError(addError.message);
    }
  }

  async function updateUser(username, changes) {
    setError("");
    setMessage("");
    try {
      await apiJson(`/api/users/${encodeURIComponent(username)}`, {
        method: "PATCH",
        body: JSON.stringify(changes),
      });
      setMessage("User updated.");
      await loadUsers();
    } catch (updateError) {
      setError(updateError.message);
    }
  }

  async function deleteUser(username) {
    setError("");
    setMessage("");
    try {
      await apiJson(`/api/users/${encodeURIComponent(username)}`, { method: "DELETE" });
      setMessage("User deleted.");
      await loadUsers();
    } catch (deleteError) {
      setError(deleteError.message);
    }
  }

  return (
    <section className="users-page">
      <form className="user-form" onSubmit={addUser}>
        <label>
          <span>Username</span>
          <input value={form.username} onChange={(event) => setForm({ ...form, username: event.target.value })} />
        </label>
        <label>
          <span>Password</span>
          <input
            type="password"
            value={form.password}
            onChange={(event) => setForm({ ...form, password: event.target.value })}
            autoComplete="new-password"
          />
        </label>
        <label>
          <span>Role</span>
          <select
            value={form.role}
            onChange={(event) => {
              const role = event.target.value;
              setForm({ ...form, role, permissions: [...DEFAULT_PERMISSIONS_BY_ROLE[role]] });
            }}
          >
            <option value="viewer">Viewer</option>
            <option value="admin">Admin</option>
          </select>
        </label>
        <button className="primary" type="submit">Add user</button>
        <PermissionChecklist
          title="Menu access"
          permissions={form.permissions}
          onChange={(permissions) => setForm({ ...form, permissions })}
        />
      </form>

      {error && <div className="notice danger">{error}</div>}
      {message && <div className="notice">{message}</div>}

      <div className="users-list">
        {users.map((user) => (
          <UserRow
            key={user.username}
            user={user}
            currentUser={currentUser}
            onSave={(changes) => updateUser(user.username, changes)}
            onDelete={() => deleteUser(user.username)}
          />
        ))}
      </div>
    </section>
  );
}

function UserRow({ user, currentUser, onSave, onDelete }) {
  const [password, setPassword] = useState("");
  const [role, setRole] = useState(user.role);
  const [permissions, setPermissions] = useState(normalizePermissions(user.role, user.permissions));

  useEffect(() => {
    setRole(user.role);
    setPermissions(normalizePermissions(user.role, user.permissions));
  }, [user.permissions, user.role]);

  return (
    <article className="user-row">
      <div>
        <strong>{user.username}</strong>
        <span>{user.role} | created {formatTimestamp(user.created_at)}</span>
      </div>

      <div className="user-row-actions">
        <label>
          <span>Role</span>
          <select
            value={role}
            onChange={(event) => {
              const nextRole = event.target.value;
              setRole(nextRole);
              setPermissions([...DEFAULT_PERMISSIONS_BY_ROLE[nextRole]]);
            }}
          >
            <option value="viewer">Viewer</option>
            <option value="admin">Admin</option>
          </select>
        </label>
        <button onClick={() => onSave({ role, permissions })}>Save access</button>
      </div>

      <PermissionChecklist title="Allowed menus" permissions={permissions} onChange={setPermissions} />

      <form
        className="password-form"
        onSubmit={(event) => {
          event.preventDefault();
          if (password) {
            onSave({ password });
            setPassword("");
          }
        }}
      >
        <input
          type="password"
          value={password}
          onChange={(event) => setPassword(event.target.value)}
          placeholder="New password"
          autoComplete="new-password"
        />
        <button type="submit">Set password</button>
      </form>

      <button disabled={user.username === currentUser.username} onClick={onDelete}>Delete</button>
    </article>
  );
}

function PermissionChecklist({ title, permissions, onChange }) {
  return (
    <fieldset className="permission-grid">
      <legend>{title}</legend>
      {ALL_PERMISSIONS.map((permission) => (
        <label key={permission} className="permission-option">
          <input
            type="checkbox"
            checked={permissions.includes(permission)}
            onChange={(event) => {
              if (event.target.checked) {
                onChange([...permissions, permission]);
                return;
              }
              onChange(permissions.filter((item) => item !== permission));
            }}
          />
          <span>{permission}</span>
        </label>
      ))}
    </fieldset>
  );
}

function fustTileLabel(tab) {
  return {
    in: "IN",
    out: "OUT",
    overview: "Overview",
    "last-actions": "Last actions",
    control: "Fust Controle",
    manage: "Fust Beheer",
    import: "Fust Import",
  }[tab] || tab.toUpperCase();
}

function FustPage({ currentUser, menuVersion }) {
  const canManageFust = hasPermission(currentUser, PERMISSIONS.FUST_MANAGE);
  const visibleTabs = [
    hasPermission(currentUser, PERMISSIONS.FUST_IN) ? "in" : null,
    hasPermission(currentUser, PERMISSIONS.FUST_OUT) ? "out" : null,
    hasPermission(currentUser, PERMISSIONS.FUST_OVERVIEW) ? "overview" : null,
    hasPermission(currentUser, PERMISSIONS.FUST_OVERVIEW) ? "last-actions" : null,
    hasPermission(currentUser, PERMISSIONS.FUST_OVERVIEW) ? "control" : null,
    canManageFust ? "manage" : null,
    canManageFust ? "import" : null,
  ].filter(Boolean);
  const [activeTab, setActiveTab] = useState("");
  const { loading: metaLoading, data: metaData, error: metaError } = useFustMeta(Boolean(currentUser));
  const { loading: actionsLoading, data: actionsData, error: actionsError, refresh } = useFustActions(Boolean(currentUser));

  useEffect(() => {
    setActiveTab("");
  }, [menuVersion]);

  useEffect(() => {
    if (activeTab && !visibleTabs.includes(activeTab)) {
      setActiveTab("");
    }
  }, [activeTab, visibleTabs]);

  if (!visibleTabs.length) {
    return <div className="notice">This account can open Fust but does not yet have an IN, OUT, Overview, or Beheer action assigned.</div>;
  }

  return (
    <section className="fust-page">
      {metaError && <div className="notice danger">Unable to load Fust master data: {metaError}</div>}
      {actionsError && <div className="notice danger">Unable to load Fust actions: {actionsError}</div>}

      {!activeTab && (
        <div className="fust-tile-grid" aria-label="Fust menu">
          {visibleTabs.map((tab) => (
            <button
              key={tab}
              type="button"
              className="fust-tile"
              onClick={() => setActiveTab(tab)}
            >
              <strong>{fustTileLabel(tab)}</strong>
            </button>
          ))}
        </div>
      )}

      {activeTab && (
        <div className="section-header fust-section-nav">
          <h2>{fustTileLabel(activeTab)}</h2>
          <button type="button" onClick={() => setActiveTab("")}>Fust menu</button>
        </div>
      )}

      {activeTab === "in" && (
        <FustActionForm
          type="IN"
          metaData={metaData}
          loading={metaLoading}
          onSaved={refresh}
        />
      )}

      {activeTab === "out" && (
        <FustActionForm
          type="OUT"
          metaData={metaData}
          loading={metaLoading}
          onSaved={refresh}
        />
      )}

      {activeTab === "overview" && (
        <FustOverview
          loading={actionsLoading}
          actions={actionsData?.actions || []}
          overview={actionsData?.overview || []}
          sourceDebug={actionsData?.source_debug || null}
          onRefresh={refresh}
        />
      )}

      {activeTab === "last-actions" && (
        <FustLastActions
          loading={actionsLoading}
          actions={actionsData?.actions || []}
          onRefresh={refresh}
        />
      )}

      {activeTab === "control" && (
        <FustControle
          loading={actionsLoading}
          actions={actionsData?.actions || []}
          onRefresh={refresh}
        />
      )}

      {activeTab === "manage" && (
        <FustBeheer
          loading={actionsLoading}
          actions={actionsData?.actions || []}
          onRefresh={refresh}
        />
      )}

      {activeTab === "import" && (
        <FustImportPanel onSaved={refresh} />
      )}
    </section>
  );
}

function FustActionForm({ type, metaData, loading, onSaved }) {
  const [form, setForm] = useState({
    action_date: new Date().toISOString().slice(0, 10),
    country: "",
    customer_name: "",
    customer_code: "",
    connect_name: "",
    remark: "",
    fustbon_reference: "",
    fustfactuur_reference: "",
    metrics: emptyFustMetrics(),
  });
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);
  const [documentFile, setDocumentFile] = useState(null);
  const [documentSkipped, setDocumentSkipped] = useState(false);
  const [documentInputKey, setDocumentInputKey] = useState(0);
  const records = metaData?.records || [];
  const countries = metaData?.countries || [];
  const customerOptions = records.filter((record) => record.country === form.country);
  const customerNames = [...new Set(customerOptions.map((record) => record.customer_name))].sort((left, right) => left.localeCompare(right));
  const connectOptions = customerOptions.filter((record) => record.customer_name === form.customer_name);
  const documentLabel = type === "IN" ? "Fustbon" : "CMR";

  function resetActionEntry() {
    setForm((current) => ({
      ...current,
      remark: "",
      fustbon_reference: "",
      fustfactuur_reference: "",
      metrics: emptyFustMetrics(),
    }));
  }

  useEffect(() => {
    if (countries.length && !form.country) {
      setForm((current) => ({ ...current, country: countries[0] }));
    }
  }, [countries, form.country]);

  useEffect(() => {
    if (customerNames.length && !customerNames.includes(form.customer_name)) {
      setForm((current) => ({
        ...current,
        customer_name: customerNames[0] || "",
        connect_name: "",
        customer_code: "",
      }));
    }
  }, [customerNames, form.customer_name]);

  useEffect(() => {
    const activeConnect = connectOptions.find((record) => record.connect_name === form.connect_name);
    if (!activeConnect && connectOptions.length) {
      const first = connectOptions[0];
      setForm((current) => ({
        ...current,
        connect_name: first.connect_name,
        customer_code: first.customer_code,
      }));
    } else if (activeConnect) {
      setForm((current) => ({
        ...current,
        customer_code: activeConnect.customer_code,
      }));
    }
  }, [connectOptions, form.connect_name]);

  async function submit(event) {
    event.preventDefault();
    if (!documentFile && !documentSkipped) {
      setError(`Choose a ${documentLabel} file or mark No ${documentLabel}.`);
      return;
    }
    setSaving(true);
    setMessage("");
    setError("");
    try {
      const document = documentSkipped
        ? { mode: "skip" }
        : {
          mode: "upload",
          file: {
            name: documentFile.name,
            type: documentFile.type || "application/octet-stream",
            content_base64: await fileToBase64(documentFile),
          },
        };
      const payload = await apiJson("/api/fust/submit", {
        method: "POST",
        body: JSON.stringify({ ...form, type, document }),
      });
      const savedDocument = type === "IN" ? payload.action.fustbon : payload.action.cmr;
      const documentStatus = savedDocument?.status === "uploaded"
        ? `${documentLabel}: uploaded`
        : savedDocument?.status === "skipped"
          ? `${documentLabel}: skipped`
          : savedDocument?.status === "failed"
            ? `${documentLabel}: ${savedDocument.error || "upload failed"}`
            : `${documentLabel}: missing`;
      setMessage(
        `${type} saved. Sheet sync: ${payload.action.sheet_sync.ok ? "ok" : payload.action.sheet_sync.error}. Email: ${payload.action.email_sync.ok ? "ok" : payload.action.email_sync.error}. ${documentStatus}`,
      );
      setDocumentFile(null);
      setDocumentSkipped(false);
      setDocumentInputKey((current) => current + 1);
      resetActionEntry();
      onSaved();
    } catch (submitError) {
      setError(submitError.message);
    } finally {
      setSaving(false);
    }
  }

  return (
    <div className="fust-form-layout">
      <form className="fust-form" onSubmit={submit}>
        <div className="form-grid">
          <label>
            <span>Date</span>
            <input
              type="date"
              value={form.action_date}
              onChange={(event) => setForm({ ...form, action_date: event.target.value })}
            />
          </label>
          <label>
            <span>Country</span>
            <select
              value={form.country}
              onChange={(event) => setForm({
                ...form,
                country: event.target.value,
                customer_name: "",
                connect_name: "",
                customer_code: "",
              })}
            >
              <option value="">Choose country</option>
              {countries.map((country) => <option key={country} value={country}>{country}</option>)}
            </select>
          </label>
          <label>
            <span>Klantnaam / carrier</span>
            <select
              value={form.customer_name}
              onChange={(event) => setForm({
                ...form,
                customer_name: event.target.value,
                connect_name: "",
                customer_code: "",
              })}
              disabled={!form.country}
            >
              <option value="">Choose klantnaam</option>
              {customerNames.map((name) => <option key={name} value={name}>{name}</option>)}
            </select>
          </label>
          <label>
            <span>Connect</span>
            <select
              value={form.connect_name}
              onChange={(event) => {
                const selected = connectOptions.find((record) => record.connect_name === event.target.value);
                setForm({
                  ...form,
                  connect_name: event.target.value,
                  customer_code: selected?.customer_code || "",
                });
              }}
              disabled={!form.customer_name}
            >
              <option value="">Choose connect</option>
              {connectOptions.map((option) => (
                <option key={`${option.customer_name}-${option.connect_name}`} value={option.connect_name}>
                  {option.connect_name}
                </option>
              ))}
            </select>
          </label>
          <label className="wide">
            <span>Remark</span>
            <input
              value={form.remark}
              onChange={(event) => setForm({ ...form, remark: event.target.value })}
              placeholder="alleen fusten"
            />
          </label>
          <label>
            <span>Fustbon</span>
            <input
              value={form.fustbon_reference}
              onChange={(event) => setForm({ ...form, fustbon_reference: event.target.value })}
              placeholder="Fustbon nummer"
            />
          </label>
          <label>
            <span>Fustfactuur</span>
            <input
              value={form.fustfactuur_reference}
              onChange={(event) => setForm({ ...form, fustfactuur_reference: event.target.value })}
              placeholder="Fustfactuur nummer"
            />
          </label>
        </div>

        <div className="metrics-grid">
          {Object.entries(form.metrics).map(([key, value]) => (
            <label key={key}>
              <span>{key.toUpperCase()}</span>
              <input
                type="number"
                min="0"
                value={value || ""}
                onChange={(event) => setForm({
                  ...form,
                  metrics: {
                    ...form.metrics,
                    [key]: Number(event.target.value || 0),
                  },
                })}
              />
            </label>
          ))}
        </div>

        {loading && <div className="notice">Loading Data sheet options...</div>}
        {message && <div className="notice">{message}</div>}
        {error && <div className="notice danger">{error}</div>}

        <div className="cmr-panel">
          <h3>{documentLabel}</h3>
          <input
            key={documentInputKey}
            type="file"
            accept="image/*,.pdf"
            capture="environment"
            disabled={documentSkipped || saving}
            onChange={(event) => {
              const nextFile = event.target.files?.[0] || null;
              setDocumentFile(nextFile);
              if (nextFile) {
                setDocumentSkipped(false);
              }
            }}
          />
          <label className="checkbox-row">
            <input
              type="checkbox"
              checked={documentSkipped}
              disabled={saving}
              onChange={(event) => {
                setDocumentSkipped(event.target.checked);
                if (event.target.checked) {
                  setDocumentFile(null);
                  setDocumentInputKey((current) => current + 1);
                }
              }}
            />
            <span>No {documentLabel}</span>
          </label>
        </div>

        <button className="primary" type="submit" disabled={saving || loading}>
          {saving ? `Saving ${type}...` : `Save ${type}`}
        </button>
      </form>
    </div>
  );
}

function FustOverview({ loading, actions, overview, sourceDebug, onRefresh }) {
  const [selectedWeek, setSelectedWeek] = useState("");
  const [selectedFromDate, setSelectedFromDate] = useState("");
  const [selectedToDate, setSelectedToDate] = useState("");
  const [selectedCountry, setSelectedCountry] = useState("");
  const [selectedCustomer, setSelectedCustomer] = useState("");
  const [expandedCountryWeek, setExpandedCountryWeek] = useState(false);
  const [showTransactionRecords, setShowTransactionRecords] = useState(false);

  function emptyTotals() {
    return {
      out: { dc: 0, cctag: 0, dcs: 0, dco: 0, pal: 0, vk: 0 },
      in: { dc: 0, cctag: 0, dcs: 0, dco: 0, pal: 0, vk: 0 },
      balance: { dc: 0, cctag: 0, dcs: 0, dco: 0, pal: 0, vk: 0 },
    };
  }

  function sumOverviewEntries(entries) {
    return entries.reduce((sum, entry) => ({
      out: {
        dc: sum.out.dc + entry.out.dc,
        cctag: sum.out.cctag + entry.out.cctag,
        dcs: sum.out.dcs + entry.out.dcs,
        dco: sum.out.dco + entry.out.dco,
        pal: sum.out.pal + entry.out.pal,
        vk: sum.out.vk + entry.out.vk,
      },
      in: {
        dc: sum.in.dc + entry.in.dc,
        cctag: sum.in.cctag + entry.in.cctag,
        dcs: sum.in.dcs + entry.in.dcs,
        dco: sum.in.dco + entry.in.dco,
        pal: sum.in.pal + entry.in.pal,
        vk: sum.in.vk + entry.in.vk,
      },
      balance: {
        dc: sum.balance.dc + entry.balance.dc,
        cctag: sum.balance.cctag + entry.balance.cctag,
        dcs: sum.balance.dcs + entry.balance.dcs,
        dco: sum.balance.dco + entry.balance.dco,
        pal: sum.balance.pal + entry.balance.pal,
        vk: sum.balance.vk + entry.balance.vk,
      },
    }), emptyTotals());
  }

  function buildOverviewFromActions(items) {
    const grouped = new Map();
    for (const action of items) {
      const key = `${action.week || ""}__${action.country}__${action.customer_name}`;
      if (!grouped.has(key)) {
        grouped.set(key, {
          week: action.week || "",
          country: action.country,
          customer_name: action.customer_name,
          out: { dc: 0, cctag: 0, dcs: 0, dco: 0, pal: 0, vk: 0 },
          in: { dc: 0, cctag: 0, dcs: 0, dco: 0, pal: 0, vk: 0 },
          balance: { dc: 0, cctag: 0, dcs: 0, dco: 0, pal: 0, vk: 0 },
        });
      }
      const entry = grouped.get(key);
      const target = action.type === "OUT" ? entry.out : entry.in;
      target.dc += Number(action.metrics?.dc || 0);
      target.cctag += Number(action.metrics?.cctag || 0);
      target.dcs += Number(action.metrics?.dcs || 0);
      target.dco += Number(action.metrics?.dco || 0);
      target.pal += Number(action.metrics?.pal || 0);
      target.vk += Number(action.metrics?.vk || 0);
      entry.balance = {
        dc: entry.in.dc - entry.out.dc,
        cctag: entry.in.cctag - entry.out.cctag,
        dcs: entry.in.dcs - entry.out.dcs,
        dco: entry.in.dco - entry.out.dco,
        pal: entry.in.pal - entry.out.pal,
        vk: entry.in.vk - entry.out.vk,
      };
    }

    return [...grouped.values()].sort((left, right) => (
      String(left.customer_name).localeCompare(String(right.customer_name))
    ));
  }

  function downloadExcelFriendlyTable(filename, headers, rows) {
    const tsv = [
      headers.join("\t"),
      ...rows.map((row) => row.map((value) => String(value ?? "").replaceAll("\t", " ")).join("\t")),
    ].join("\n");
    const blob = new Blob([tsv], { type: "text/tab-separated-values;charset=utf-8;" });
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.URL.revokeObjectURL(url);
  }

  const weekOptions = [...new Set(actions.map((action) => String(action.week || "")).filter(Boolean))]
    .sort((left, right) => Number(right) - Number(left));
  const datedActions = actions.filter((action) => {
    const actionDate = String(action.action_date || "");
    if (selectedFromDate && actionDate && actionDate < selectedFromDate) {
      return false;
    }
    if (selectedToDate && actionDate && actionDate > selectedToDate) {
      return false;
    }
    return true;
  });
  const scopedActions = datedActions.filter((action) => !selectedWeek || String(action.week || "") === selectedWeek);
  const weekFilteredOverview = buildOverviewFromActions(scopedActions);
  const countryOptions = [...new Set(weekFilteredOverview.map((entry) => entry.country).filter(Boolean))]
    .sort((left, right) => left.localeCompare(right));
  const countryFilteredOverview = weekFilteredOverview.filter((entry) => !selectedCountry || entry.country === selectedCountry);
  const customerOptions = [...new Set(countryFilteredOverview.map((entry) => entry.customer_name).filter(Boolean))]
    .sort((left, right) => left.localeCompare(right));

  useEffect(() => {
    if (selectedCountry && !countryOptions.includes(selectedCountry)) {
      setSelectedCountry("");
    }
  }, [countryOptions, selectedCountry]);

  useEffect(() => {
    if (selectedCustomer && !customerOptions.includes(selectedCustomer)) {
      setSelectedCustomer("");
      setShowTransactionRecords(false);
    }
  }, [customerOptions, selectedCustomer]);

  const countryRows = useMemo(() => {
    const grouped = new Map();
    for (const entry of weekFilteredOverview) {
      if (!grouped.has(entry.country)) {
        grouped.set(entry.country, []);
      }
      grouped.get(entry.country).push(entry);
    }

    return [...grouped.entries()]
      .map(([country, entries]) => ({
        week: selectedWeek || "All weeks",
        country,
        customer_name: `${entries.length} cust/transports`,
        ...sumOverviewEntries(entries),
        row_count: entries.length,
      }))
      .sort((left, right) => left.country.localeCompare(right.country));
  }, [selectedWeek, weekFilteredOverview]);

  const weekRows = useMemo(() => {
    const grouped = new Map();
    for (const entry of countryFilteredOverview) {
      const week = String(entry.week || "");
      if (!grouped.has(week)) {
        grouped.set(week, []);
      }
      grouped.get(week).push(entry);
    }

    return [...grouped.entries()]
      .map(([week, entries]) => ({
        week,
        country: selectedCountry,
        customer_name: `${entries.length} cust/transports`,
        ...sumOverviewEntries(entries),
        row_count: entries.length,
      }))
      .sort((left, right) => Number(right.week || 0) - Number(left.week || 0));
  }, [countryFilteredOverview, selectedCountry]);

  const customerRows = countryFilteredOverview
    .sort((left, right) => left.customer_name.localeCompare(right.customer_name));
  const collapsedCountryWeekRows = selectedCountry && selectedWeek ? weekRows : [];
  const scopedOverview = selectedCountry
    ? (selectedWeek && !expandedCountryWeek ? collapsedCountryWeekRows : (selectedWeek ? customerRows : weekRows))
    : countryRows;
  const visibleOverview = selectedCustomer
    ? customerRows.filter((entry) => entry.customer_name === selectedCustomer)
    : scopedOverview;
  const totals = sumOverviewEntries(visibleOverview);
  const transactionRecords = datedActions
    .filter((action) => showTransactionRecords)
    .filter((action) => selectedWeek && String(action.week || "") === selectedWeek)
    .filter((action) => selectedCountry && action.country === selectedCountry)
    .filter((action) => selectedCustomer && action.customer_name === selectedCustomer)
    .sort((left, right) => {
      const rightDate = String(right.action_date || right.created_at || "");
      const leftDate = String(left.action_date || left.created_at || "");
      return rightDate.localeCompare(leftDate);
    });
  const exportRows = visibleOverview.map((entry, index) => [
    selectedCountry && selectedWeek && !selectedCustomer && index > 0 ? "" : (entry.week || ""),
    selectedCountry && selectedWeek && !selectedCustomer && index > 0 ? "" : entry.country,
    entry.customer_name,
    entry.out.dc,
    entry.in.dc,
    entry.balance.dc,
    entry.out.cctag,
    entry.in.cctag,
    entry.balance.cctag,
    entry.out.dcs,
    entry.in.dcs,
    entry.balance.dcs,
    entry.out.dco,
    entry.in.dco,
    entry.balance.dco,
    entry.out.vk,
    entry.in.vk,
    entry.balance.vk,
    entry.out.pal,
  ]);

  if (loading) {
    return <div className="notice">Loading Fust overview...</div>;
  }

  return (
    <div className="overview-stack">
      <div className="data-table-card">
        <div className="overview-filters">
          <label>
            <span>Week</span>
            <select
              value={selectedWeek}
              onChange={(event) => {
                setSelectedWeek(event.target.value);
                setSelectedCustomer("");
                setExpandedCountryWeek(false);
                setShowTransactionRecords(false);
              }}
            >
              <option value="">All weeks</option>
              {weekOptions.map((week) => <option key={week} value={week}>{week}</option>)}
            </select>
          </label>
          <label>
            <span>From date</span>
            <input
              type="date"
              value={selectedFromDate}
              onChange={(event) => {
                setSelectedFromDate(event.target.value);
                setExpandedCountryWeek(false);
                setShowTransactionRecords(false);
              }}
            />
          </label>
          <label>
            <span>To date</span>
            <input
              type="date"
              value={selectedToDate}
              onChange={(event) => {
                setSelectedToDate(event.target.value);
                setExpandedCountryWeek(false);
                setShowTransactionRecords(false);
              }}
            />
          </label>
          <label>
            <span>Country</span>
            <select
              value={selectedCountry}
              onChange={(event) => {
                setSelectedCountry(event.target.value);
                setSelectedCustomer("");
                setExpandedCountryWeek(false);
                setShowTransactionRecords(false);
              }}
            >
              <option value="">All countries</option>
              {countryOptions.map((country) => <option key={country} value={country}>{country}</option>)}
            </select>
          </label>
          <label>
            <span>Cust/transport</span>
            <select
              value={selectedCustomer}
              onChange={(event) => {
                setSelectedCustomer(event.target.value);
                setExpandedCountryWeek(false);
                setShowTransactionRecords(false);
              }}
              disabled={!selectedCountry}
            >
              <option value="">All cust/transports</option>
              {customerOptions.map((customer) => <option key={customer} value={customer}>{customer}</option>)}
            </select>
          </label>
        </div>
        <div className="section-header">
          <h2>
            {!selectedCountry && "Country totals"}
            {selectedCountry && !selectedWeek && "Week totals"}
            {selectedCountry && selectedWeek && !expandedCountryWeek && !selectedCustomer && "Country total"}
            {selectedCountry && selectedWeek && (expandedCountryWeek || selectedCustomer) && "Customer totals"}
          </h2>
          <button
            type="button"
            onClick={() => downloadExcelFriendlyTable(
              `fust-overview-${selectedWeek || "all-weeks"}-${selectedCountry || "all-countries"}-${selectedCustomer || "active-table"}.tsv`,
              ["Week", "Country", "Cust/transport", "DC out", "DC in", "DC balance", "CCTag out", "CCTag in", "CCTag balance", "DCS out", "DCS in", "DCS balance", "DCO out", "DCO in", "DCO balance", "VK out", "VK in", "VK balance", "pal out"],
              exportRows,
            )}
            disabled={!visibleOverview.length}
          >
            Export active table
          </button>
        </div>
        <div className="table-wrap">
          <table className="data-table balance-table">
            <thead>
              <tr>
                <th>Week</th>
                <th>Country</th>
                <th>Cust/transport</th>
                <th>DC out</th>
                <th>DC in</th>
                <th>DC balance</th>
                <th>CCTag out</th>
                <th>CCTag in</th>
                <th>CCTag balance</th>
                <th>DCS out</th>
                <th>DCS in</th>
                <th>DCS balance</th>
                <th>DCO out</th>
                <th>DCO in</th>
                <th>DCO balance</th>
                <th>VK out</th>
                <th>VK in</th>
                <th>VK balance</th>
                <th>pal out</th>
              </tr>
            </thead>
            <tbody>
              {visibleOverview.map((entry) => (
                <tr
                  key={`${entry.week}-${entry.country}-${entry.customer_name}`}
                  className={selectedCustomer === entry.customer_name ? "clickable-row active-row" : "clickable-row"}
                  onClick={() => {
                    if (!selectedCountry) {
                      setSelectedCountry(entry.country);
                      setSelectedCustomer("");
                      setShowTransactionRecords(false);
                      return;
                    }
                    if (selectedCountry && !selectedWeek) {
                      setSelectedWeek(String(entry.week || ""));
                      setSelectedCustomer("");
                      setExpandedCountryWeek(true);
                      setShowTransactionRecords(false);
                      return;
                    }
                    if (selectedCountry && selectedWeek && !expandedCountryWeek && !selectedCustomer) {
                      setExpandedCountryWeek(true);
                      setShowTransactionRecords(false);
                      return;
                    }
                    if (selectedCountry && selectedWeek && selectedCustomer) {
                      setShowTransactionRecords((current) => !current);
                      return;
                    }
                    setSelectedCustomer(entry.customer_name);
                    setShowTransactionRecords(false);
                  }}
                >
                  <td>{entry.week || ""}</td>
                  <td>{entry.country}</td>
                  <td>{entry.customer_name}</td>
                  <td>{entry.out.dc}</td>
                  <td>{entry.in.dc}</td>
                  <td>{entry.balance.dc}</td>
                  <td>{entry.out.cctag}</td>
                  <td>{entry.in.cctag}</td>
                  <td>{entry.balance.cctag}</td>
                  <td>{entry.out.dcs}</td>
                  <td>{entry.in.dcs}</td>
                  <td>{entry.balance.dcs}</td>
                  <td>{entry.out.dco}</td>
                  <td>{entry.in.dco}</td>
                  <td>{entry.balance.dco}</td>
                  <td>{entry.out.vk}</td>
                  <td>{entry.in.vk}</td>
                  <td>{entry.balance.vk}</td>
                  <td>{entry.out.pal}</td>
                </tr>
              ))}
              {!!visibleOverview.length && (
                <tr className="summary-row">
                  <td colSpan="3"><strong>Total</strong></td>
                  <td><strong>{totals.out.dc}</strong></td>
                  <td><strong>{totals.in.dc}</strong></td>
                  <td><strong>{totals.balance.dc}</strong></td>
                  <td><strong>{totals.out.cctag}</strong></td>
                  <td><strong>{totals.in.cctag}</strong></td>
                  <td><strong>{totals.balance.cctag}</strong></td>
                  <td><strong>{totals.out.dcs}</strong></td>
                  <td><strong>{totals.in.dcs}</strong></td>
                  <td><strong>{totals.balance.dcs}</strong></td>
                  <td><strong>{totals.out.dco}</strong></td>
                  <td><strong>{totals.in.dco}</strong></td>
                  <td><strong>{totals.balance.dco}</strong></td>
                  <td><strong>{totals.out.vk}</strong></td>
                  <td><strong>{totals.in.vk}</strong></td>
                  <td><strong>{totals.balance.vk}</strong></td>
                  <td><strong>{totals.out.pal}</strong></td>
                </tr>
              )}
              {!visibleOverview.length && (
                <tr>
                  <td colSpan="19">No overview rows found for the selected filters.</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
        {showTransactionRecords && selectedWeek && selectedCountry && selectedCustomer && (
          <div className="transaction-detail">
            <div className="section-header">
              <h3>Transactions</h3>
              <button type="button" onClick={() => setShowTransactionRecords(false)}>Close</button>
            </div>
            <div className="table-wrap">
              <table className="data-table action-table">
                <thead>
                  <tr>
                    <th>Type</th>
                    <th>Date</th>
                    <th>Country</th>
                    <th>Cust/transport</th>
                    <th>Connect</th>
                    <th>DC</th>
                    <th>CCTag</th>
                    <th>DCS</th>
                    <th>DCO</th>
                    <th>PAL</th>
                    <th>VK</th>
                    <th>Document</th>
                    <th>Remark</th>
                  </tr>
                </thead>
                <tbody>
                  {transactionRecords.map((action) => (
                    <tr key={action.id}>
                      <td>{action.type}</td>
                      <td>{action.action_date}</td>
                      <td>{action.country}</td>
                      <td>{action.customer_name}</td>
                      <td>{action.connect_name}</td>
                      <td>{action.metrics?.dc || 0}</td>
                      <td>{action.metrics?.cctag || 0}</td>
                      <td>{action.metrics?.dcs || 0}</td>
                      <td>{action.metrics?.dco || 0}</td>
                      <td>{action.metrics?.pal || 0}</td>
                      <td>{action.metrics?.vk || 0}</td>
                      <td><DocumentStatus action={action} /></td>
                      <td>{action.remark || "-"}</td>
                    </tr>
                  ))}
                  {!transactionRecords.length && (
                    <tr>
                      <td colSpan="13">No transactions were found for this filter.</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

function yesterdayIso() {
  const now = new Date();
  const yesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
  return `${yesterday.getFullYear()}-${String(yesterday.getMonth() + 1).padStart(2, "0")}-${String(yesterday.getDate()).padStart(2, "0")}`;
}

function isFustActionConfirmed(action) {
  return Boolean(String(action?.confirmed_at || "").trim());
}

function FustActionTable({
  loading,
  actions,
  onRefresh,
  title,
  readOnly = false,
  allowConfirm = false,
  allowManage = false,
  defaultDate = "",
  unconfirmedOnly = false,
  emptyMessage = "No actions were found.",
}) {
  const [busyActionId, setBusyActionId] = useState("");
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const [editingActionId, setEditingActionId] = useState("");
  const [editForm, setEditForm] = useState(null);
  const [typeFilter, setTypeFilter] = useState("");
  const [dateFilter, setDateFilter] = useState(defaultDate);
  const [countryFilter, setCountryFilter] = useState("");
  const [customerFilter, setCustomerFilter] = useState("");

  useEffect(() => {
    setDateFilter(defaultDate);
  }, [defaultDate]);

  function startEdit(action) {
    setEditingActionId(action.id);
    setEditForm({
      type: action.type,
      action_date: action.action_date || "",
      country: action.country || "",
      customer_name: action.customer_name || "",
      connect_name: action.connect_name || "",
      customer_code: action.customer_code || "",
      remark: action.remark || "",
      fustbon_reference: action.fustbon_reference || "",
      fustfactuur_reference: action.fustfactuur_reference || "",
      metrics: {
        dc: Number(action.metrics?.dc || 0),
        cctag: Number(action.metrics?.cctag || 0),
        dcs: Number(action.metrics?.dcs || 0),
        dco: Number(action.metrics?.dco || 0),
        pal: Number(action.metrics?.pal || 0),
        vk: Number(action.metrics?.vk || 0),
      },
    });
  }

  function cancelEdit() {
    setEditingActionId("");
    setEditForm(null);
  }

  async function saveEdit(actionId) {
    setBusyActionId(`${actionId}:save`);
    setMessage("");
    setError("");
    try {
      await apiJson(`/api/fust/actions/${encodeURIComponent(actionId)}`, {
        method: "PUT",
        body: JSON.stringify(editForm),
      });
      setMessage("Action updated.");
      cancelEdit();
      onRefresh();
    } catch (saveError) {
      setError(saveError.message);
    } finally {
      setBusyActionId("");
    }
  }

  async function retryAction(actionId, kind) {
    setBusyActionId(`${actionId}:${kind}`);
    setMessage("");
    setError("");
    try {
      await apiJson(`/api/fust/actions/${encodeURIComponent(actionId)}/${kind}`, {
        method: "POST",
      });
      setMessage(kind === "retry-sheet" ? "Sheet sync retried." : "Email resend retried.");
      onRefresh();
    } catch (retryError) {
      setError(retryError.message);
    } finally {
      setBusyActionId("");
    }
  }

  async function deleteLocalAction(actionId) {
    setBusyActionId(`${actionId}:delete`);
    setMessage("");
    setError("");
    try {
      await apiJson(`/api/fust/actions/${encodeURIComponent(actionId)}`, {
        method: "DELETE",
      });
      setMessage("Action deleted.");
      onRefresh();
    } catch (deleteError) {
      setError(deleteError.message);
    } finally {
      setBusyActionId("");
    }
  }

  async function toggleConfirm(actionId, confirmed) {
    setBusyActionId(`${actionId}:${confirmed ? "unconfirm" : "confirm"}`);
    setMessage("");
    setError("");
    try {
      await apiJson(`/api/fust/actions/${encodeURIComponent(actionId)}/${confirmed ? "unconfirm" : "confirm"}`, {
        method: "POST",
      });
      setMessage(confirmed ? "Confirmation removed." : "Action confirmed.");
      onRefresh();
    } catch (confirmError) {
      setError(confirmError.message);
    } finally {
      setBusyActionId("");
    }
  }

  const typeOptions = [...new Set(actions.map((action) => action.type).filter(Boolean))].sort((left, right) => left.localeCompare(right));
  const countryOptions = [...new Set(actions.map((action) => action.country).filter(Boolean))].sort((left, right) => left.localeCompare(right));
  const customerOptions = [...new Set(actions.map((action) => action.customer_name).filter(Boolean))].sort((left, right) => left.localeCompare(right));
  const visibleActions = actions
    .filter((action) => !unconfirmedOnly || !isFustActionConfirmed(action))
    .filter((action) => !typeFilter || action.type === typeFilter)
    .filter((action) => !dateFilter || String(action.action_date || "") === dateFilter)
    .filter((action) => !countryFilter || action.country === countryFilter)
    .filter((action) => !customerFilter || action.customer_name === customerFilter);

  if (loading) {
    return <div className="notice">Loading Fust actions...</div>;
  }

  return (
    <div className="overview-stack">
      {message && <div className="notice">{message}</div>}
      {error && <div className="notice danger">{error}</div>}

      <div className="data-table-card">
        <div className="section-header">
          <h2>{title}</h2>
        </div>
        <div className="overview-filters">
          <label>
            <span>Type</span>
            <select value={typeFilter} onChange={(event) => setTypeFilter(event.target.value)}>
              <option value="">All types</option>
              {typeOptions.map((type) => <option key={type} value={type}>{type}</option>)}
            </select>
          </label>
          <label>
            <span>Date</span>
            <input type="date" value={dateFilter} onChange={(event) => setDateFilter(event.target.value)} />
          </label>
          <label>
            <span>Country</span>
            <select value={countryFilter} onChange={(event) => setCountryFilter(event.target.value)}>
              <option value="">All countries</option>
              {countryOptions.map((country) => <option key={country} value={country}>{country}</option>)}
            </select>
          </label>
          <label>
            <span>Klantnaam</span>
            <select value={customerFilter} onChange={(event) => setCustomerFilter(event.target.value)}>
              <option value="">All klantnamen</option>
              {customerOptions.map((customer) => <option key={customer} value={customer}>{customer}</option>)}
            </select>
          </label>
        </div>
        <div className="table-wrap">
          <table className="data-table action-table">
            <thead>
              <tr>
                <th>Type</th>
                <th>Date</th>
                <th>Country</th>
                <th>Klantnaam</th>
                <th>Connect</th>
                <th>DC</th>
                <th>DCS</th>
                <th>DCO</th>
                <th>CCTag</th>
                <th>PAL</th>
                <th>VK</th>
                <th>Document</th>
                <th>Remark</th>
                <th>Fustbon</th>
                <th>Fustfactuur</th>
                <th>Sheet</th>
                <th>Email</th>
                <th>Confirmed</th>
                <th>Import</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {visibleActions.map((action) => {
                const isEditing = editingActionId === action.id && editForm;
                const confirmed = isFustActionConfirmed(action);
                const canModify = !readOnly && (allowManage || !confirmed);
                return (
                  <tr key={action.id}>
                    <td>{isEditing ? <select value={editForm.type} onChange={(event) => setEditForm({ ...editForm, type: event.target.value })}><option value="IN">IN</option><option value="OUT">OUT</option></select> : action.type}</td>
                    <td>{isEditing ? <input type="date" value={editForm.action_date} onChange={(event) => setEditForm({ ...editForm, action_date: event.target.value })} /> : action.action_date}</td>
                    <td>{isEditing ? <input value={editForm.country} onChange={(event) => setEditForm({ ...editForm, country: event.target.value })} /> : action.country}</td>
                    <td>{isEditing ? <input value={editForm.customer_name} onChange={(event) => setEditForm({ ...editForm, customer_name: event.target.value })} /> : action.customer_name}</td>
                    <td>{isEditing ? <input value={editForm.connect_name} onChange={(event) => setEditForm({ ...editForm, connect_name: event.target.value })} /> : action.connect_name}</td>
                    {["dc", "dcs", "dco", "cctag", "pal", "vk"].map((metric) => (
                      <td key={metric}>
                        {isEditing ? (
                          <input
                            className="metric-edit"
                            type="number"
                            min="0"
                            value={editForm.metrics[metric]}
                            onChange={(event) => setEditForm({
                              ...editForm,
                              metrics: { ...editForm.metrics, [metric]: Number(event.target.value || 0) },
                            })}
                          />
                        ) : (action.metrics?.[metric] || 0)}
                      </td>
                    ))}
                    <td><DocumentStatus action={action} /></td>
                    <td>{isEditing ? <input value={editForm.remark} onChange={(event) => setEditForm({ ...editForm, remark: event.target.value })} /> : (action.remark || "-")}</td>
                    <td>{isEditing ? <input value={editForm.fustbon_reference} onChange={(event) => setEditForm({ ...editForm, fustbon_reference: event.target.value })} /> : (action.fustbon_reference || "-")}</td>
                    <td>{isEditing ? <input value={editForm.fustfactuur_reference} onChange={(event) => setEditForm({ ...editForm, fustfactuur_reference: event.target.value })} /> : (action.fustfactuur_reference || "-")}</td>
                    <td>{action.sheet_sync?.ok ? "ok" : action.sheet_sync?.error || "-"}</td>
                    <td>{action.email_sync?.ok ? "ok" : action.email_sync?.error || "-"}</td>
                    <td>{confirmed ? `${formatTimestamp(action.confirmed_at)}${action.confirmed_by ? ` by ${action.confirmed_by}` : ""}` : "-"}</td>
                    <td>{action.import_source?.file_name ? `${action.import_source.file_name}${action.import_source.row_number ? ` row ${action.import_source.row_number}` : ""}` : "-"}</td>
                    <td>
                      <div className="retry-actions">
                        {isEditing ? (
                          <>
                            <button type="button" disabled={busyActionId === `${action.id}:save`} onClick={() => saveEdit(action.id)}>Save</button>
                            <button type="button" onClick={cancelEdit}>Cancel</button>
                          </>
                        ) : (
                          <>
                            {canModify && <button type="button" onClick={() => startEdit(action)}>Edit</button>}
                            {allowConfirm && !confirmed && <button type="button" disabled={busyActionId === `${action.id}:confirm`} onClick={() => toggleConfirm(action.id, false)}>Confirm</button>}
                            {allowManage && confirmed && <button type="button" disabled={busyActionId === `${action.id}:unconfirm`} onClick={() => toggleConfirm(action.id, true)}>Unconfirm</button>}
                          </>
                        )}
                        {!readOnly && !action.sheet_sync?.ok && <button type="button" disabled={busyActionId === `${action.id}:retry-sheet`} onClick={() => retryAction(action.id, "retry-sheet")}>Retry sheet</button>}
                        {!readOnly && !action.email_sync?.ok && <button type="button" disabled={busyActionId === `${action.id}:retry-email`} onClick={() => retryAction(action.id, "retry-email")}>Retry email</button>}
                        {canModify && !isEditing && <button type="button" disabled={busyActionId === `${action.id}:delete`} onClick={() => deleteLocalAction(action.id)}>Delete</button>}
                      </div>
                    </td>
                  </tr>
                );
              })}
              {!visibleActions.length && (
                <tr>
                  <td colSpan="20">{emptyMessage}</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function FustLastActions({ loading, actions, onRefresh }) {
  return <FustActionTable loading={loading} actions={actions} onRefresh={onRefresh} title="Last actions" readOnly emptyMessage="No actions were found." />;
}

function FustControle({ loading, actions, onRefresh }) {
  return (
    <FustActionTable
      loading={loading}
      actions={actions}
      onRefresh={onRefresh}
      title="Fust Controle"
      defaultDate={yesterdayIso()}
      unconfirmedOnly
      allowConfirm
      emptyMessage="No unconfirmed actions were found for this filter."
    />
  );
}

function FustBeheer({ loading, actions, onRefresh }) {
  return (
    <FustActionTable
      loading={loading}
      actions={actions}
      onRefresh={onRefresh}
      title="Fust Beheer"
      allowManage
      emptyMessage="No actions were found in Fust Beheer."
    />
  );
}

function FustImportPanel({ onSaved }) {
  const [file, setFile] = useState(null);
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const [preview, setPreview] = useState(null);
  const [selectedImportKeys, setSelectedImportKeys] = useState([]);

  function selectableRows(rows) {
    return (rows || []).filter((row) => row.status !== "missing_connect" && row.status !== "locked");
  }

  function syncPreviewSelection(payload) {
    const rows = payload?.rows || [];
    setPreview(payload);
    setSelectedImportKeys(selectableRows(rows).map((row) => row.import_key).filter(Boolean));
  }

  async function analyzeImport() {
    if (!file) {
      setError("Choose an import file first.");
      return;
    }
    setBusy(true);
    setMessage("");
    setError("");
    try {
      const payload = await apiJson("/api/fust/import/preview", {
        method: "POST",
        body: JSON.stringify({
          file: {
            name: file.name,
            type: file.type || "application/octet-stream",
            content_base64: await fileToBase64(file),
          },
        }),
      });
      syncPreviewSelection(payload);
      setMessage(`Preview ready: ${payload.summary?.total_rows || 0} rows found.`);
    } catch (previewError) {
      setError(previewError.message);
    } finally {
      setBusy(false);
    }
  }

  async function applyImport() {
    if (!file) {
      setError("Choose an import file first.");
      return;
    }
    setBusy(true);
    setMessage("");
    setError("");
    try {
      const payload = await apiJson("/api/fust/import/apply", {
        method: "POST",
        body: JSON.stringify({
          file: {
            name: file.name,
            type: file.type || "application/octet-stream",
            content_base64: await fileToBase64(file),
          },
          selected_import_keys: selectedImportKeys,
        }),
      });
      syncPreviewSelection({ rows: payload.rows, summary: payload.summary, file_name: payload.summary?.file_name || file.name, sheet_name: payload.summary?.sheet_name || "Overzicht" });
      setMessage(`Import done. Selected ${payload.summary?.selected_rows || 0}, created ${payload.summary?.created || 0}, updated ${payload.summary?.updated || 0}, skipped ${payload.summary?.skipped || 0}, locked ${payload.summary?.locked || 0}, missing connect ${payload.summary?.missing_connect || 0}, failed ${payload.summary?.failed || 0}.`);
      onSaved();
    } catch (importError) {
      setError(importError.message);
    } finally {
      setBusy(false);
    }
  }

  const allSelectableRows = selectableRows(preview?.rows || []);
  const allSelected = allSelectableRows.length > 0 && allSelectableRows.every((row) => selectedImportKeys.includes(row.import_key));

  function toggleImportKey(importKey, checked) {
    setSelectedImportKeys((current) => (
      checked
        ? [...new Set([...current, importKey])]
        : current.filter((value) => value !== importKey)
    ));
  }

  function toggleAll(checked) {
    setSelectedImportKeys(checked ? allSelectableRows.map((row) => row.import_key).filter(Boolean) : []);
  }

  return (
    <div className="overview-stack">
      {message && <div className="notice">{message}</div>}
      {error && <div className="notice danger">{error}</div>}
      <div className="data-table-card">
        <div className="section-header">
          <h2>Fust Import</h2>
        </div>
        <div className="form-grid">
          <label className="wide">
            <span>Overzicht file</span>
            <input type="file" accept=".xlsx,.xls,.csv" onChange={(event) => setFile(event.target.files?.[0] || null)} />
          </label>
        </div>
        <div className="row-actions hal-actions-row">
          <button type="button" onClick={analyzeImport} disabled={busy || !file}>{busy ? "Analyzing..." : "Preview import"}</button>
          <button type="button" className="primary" onClick={applyImport} disabled={busy || !file}>{busy ? "Importing..." : "Import actions"}</button>
        </div>
        {preview && (
          <>
            <div className="notice">
              {preview.file_name || file?.name || "file"} | {preview.sheet_name || "Overzicht"} | total {preview.summary?.total_rows || 0}, selected {selectedImportKeys.length}, new {preview.summary?.new_rows ?? preview.summary?.created ?? 0}, update {preview.summary?.update_rows ?? preview.summary?.updated ?? 0}, locked {preview.summary?.locked_rows ?? preview.summary?.locked ?? 0}, missing connect {preview.summary?.missing_connect_rows ?? preview.summary?.missing_connect ?? 0}
            </div>
            <div className="row-actions hal-actions-row">
              <label className="checkbox-row">
                <input type="checkbox" checked={allSelected} onChange={(event) => toggleAll(event.target.checked)} />
                <span>Select all importable rows</span>
              </label>
            </div>
            <div className="table-wrap">
              <table className="data-table action-table">
                <thead>
                  <tr>
                    <th>Use</th>
                    <th>Date</th>
                    <th>Customer</th>
                    <th>Connect</th>
                    <th>DC</th>
                    <th>DCS</th>
                    <th>DCO</th>
                    <th>Status</th>
                    <th>Note</th>
                  </tr>
                </thead>
                <tbody>
                  {(preview.rows || []).map((row, index) => (
                    <tr key={`${row.action_date}-${row.customer_name}-${index}`}>
                      <td>
                        {row.status === "missing_connect" || row.status === "locked" ? (
                          "-"
                        ) : (
                          <input
                            type="checkbox"
                            checked={selectedImportKeys.includes(row.import_key)}
                            onChange={(event) => toggleImportKey(row.import_key, event.target.checked)}
                          />
                        )}
                      </td>
                      <td>{row.action_date}</td>
                      <td>{row.customer_name}</td>
                      <td>{row.connect_name || "-"}</td>
                      <td>{row.metrics?.dc || 0}</td>
                      <td>{row.metrics?.dcs || 0}</td>
                      <td>{row.metrics?.dco || 0}</td>
                      <td>{row.status}</td>
                      <td>{row.note || "-"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

function cmrFolderLines(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return "";
  }
  return Object.entries(value)
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([country, folderId]) => `${country}=${folderId}`)
    .join("\n");
}

function parseCmrFolderLines(value) {
  return Object.fromEntries(
    String(value || "")
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean)
      .map((line) => {
        const [country, ...folderParts] = line.split("=");
        return [String(country || "").trim().toUpperCase(), folderParts.join("=").trim()];
      })
      .filter(([country, folderId]) => country && folderId),
  );
}


function todayInputValue() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function timeInputValue() {
  return new Date().toLocaleTimeString("nl-NL", { hour12: false }).slice(0, 5);
}

function employeeOptionLabel(employee) {
  return `${employee.name} (${employee.tbnr})`;
}

function workedTimeToMinutes(value) {
  const match = String(value || "").match(/^(\d+):(\d{2})$/);
  if (!match) {
    return 0;
  }
  return (Number(match[1]) * 60) + Number(match[2]);
}

function minutesToWorkedTime(minutes) {
  const rounded = Math.max(0, Math.round(minutes));
  const hours = Math.floor(rounded / 60);
  const remainder = rounded % 60;
  return `${hours}:${String(remainder).padStart(2, "0")}`;
}

function ClockPage({ currentUser, publicMode = false }) {
  const canManage = !publicMode && hasPermission(currentUser, PERMISSIONS.CLOCK_MANAGE);
  const employeeApiPath = publicMode ? "/api/public/clock/employees" : "/api/clock/employees";
  const scanApiPath = publicMode ? "/api/public/clock/scan" : "/api/clock/scan";
  const [activeTab, setActiveTab] = useState("clock");
  const [employees, setEmployees] = useState([]);
  const [records, setRecords] = useState([]);
  const [sessions, setSessions] = useState([]);
  const [selectedDate, setSelectedDate] = useState(todayInputValue());
  const [exportFrom, setExportFrom] = useState(todayInputValue());
  const [exportTo, setExportTo] = useState(todayInputValue());
  const [scanCode, setScanCode] = useState("");
  const [manual, setManual] = useState({ employeeKey: "", action_date: todayInputValue(), in_time: "", out_time: "" });
  const [editingId, setEditingId] = useState("");
  const [editForm, setEditForm] = useState({});
  const [exportEditingId, setExportEditingId] = useState("");
  const [exportEditForm, setExportEditForm] = useState({});
  const [loading, setLoading] = useState(false);
  const [employeeLoading, setEmployeeLoading] = useState(false);
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const [nowLabel, setNowLabel] = useState(timeInputValue());
  const scanInputRef = useRef(null);
  const scanCodeRef = useRef("");

  async function loadEmployees() {
    setEmployeeLoading(true);
    try {
      const payload = await apiJson(employeeApiPath);
      setEmployees(payload.employees || []);
    } catch (loadError) {
      setError(loadError.message);
    } finally {
      setEmployeeLoading(false);
    }
  }

  async function loadRecords(date = selectedDate) {
    setLoading(true);
    try {
      const payload = await apiJson(`/api/clock/records?date=${encodeURIComponent(date)}`);
      setRecords(payload.records || []);
      setSessions(payload.sessions || []);
    } catch (loadError) {
      setError(loadError.message);
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    loadEmployees();
  }, [employeeApiPath]);

  useEffect(() => {
    if (publicMode) {
      setRecords([]);
      setSessions([]);
      return;
    }
    loadRecords(selectedDate);
  }, [publicMode, selectedDate]);

  useEffect(() => {
    const interval = window.setInterval(() => setNowLabel(timeInputValue()), 1000);
    return () => window.clearInterval(interval);
  }, []);

  useEffect(() => {
    scanCodeRef.current = scanCode;
  }, [scanCode]);

  const employeeByKey = useMemo(() => {
    const map = new Map();
    for (const employee of employees) {
      const optionLabel = employeeOptionLabel(employee);
      map.set(employee.tbnr, employee);
      map.set(optionLabel, employee);
      map.set(optionLabel.toUpperCase(), employee);
    }
    return map;
  }, [employees]);

  const selectedManualEmployee = employeeByKey.get(manual.employeeKey);

  const selectedDateTotalWorked = useMemo(() => {
    const totalMinutes = sessions.reduce((total, session) => total + workedTimeToMinutes(session.row?.[6]), 0);
    return minutesToWorkedTime(totalMinutes);
  }, [sessions]);

  function focusScanInput() {
    window.setTimeout(() => scanInputRef.current?.focus(), 0);
  }

  async function processScan(codeValue = scanCodeRef.current) {
    const normalizedCode = String(codeValue || "").trim().toUpperCase();
    if (!normalizedCode) {
      setError("Scan a badge code first");
      focusScanInput();
      return;
    }
    setBusy(true);
    setError("");
    setMessage("");
    try {
      const scanDate = todayInputValue();
      const payload = await apiJson(scanApiPath, {
        method: "POST",
        body: JSON.stringify({ code: normalizedCode, action_date: scanDate, action_time: `${timeInputValue()}:00` }),
      });
      setSelectedDate(scanDate);
      setRecords(payload.records || []);
      setSessions(payload.sessions || []);
      setScanCode("");
      setMessage(`${payload.record.name}: ${payload.record.action_time} ${payload.record.direction}`);
    } catch (scanError) {
      setError(scanError.message);
    } finally {
      setBusy(false);
      focusScanInput();
    }
  }

  async function submitScan(event) {
    event.preventDefault();
    await processScan(scanCodeRef.current);
  }

  async function submitManual(event) {
    event?.preventDefault?.();
    if (!canManage) {
      return;
    }
    const employee = employeeByKey.get(manual.employeeKey);
    if (!employee) {
      setError("Choose a valid employee");
      return;
    }
    if (!manual.action_date) {
      setError("Choose a valid date");
      return;
    }
    if (!manual.in_time && !manual.out_time) {
      setError("Add an IN time, an OUT time, or both");
      return;
    }
    setBusy(true);
    setError("");
    setMessage("");
    try {
      if (manual.in_time) {
        await apiJson("/api/clock/records", {
          method: "POST",
          body: JSON.stringify({
            employee,
            action_date: manual.action_date,
            action_time: manual.in_time.length === 5 ? `${manual.in_time}:00` : manual.in_time,
            direction: "IN",
          }),
        });
      }
      if (manual.out_time) {
        await apiJson("/api/clock/records", {
          method: "POST",
          body: JSON.stringify({
            employee,
            action_date: manual.action_date,
            action_time: manual.out_time.length === 5 ? `${manual.out_time}:00` : manual.out_time,
            direction: "OUT",
          }),
        });
      }
      setSelectedDate(manual.action_date);
      await loadRecords(manual.action_date);
      setManual({ employeeKey: "", action_date: manual.action_date, in_time: "", out_time: "" });
      setMessage(`Clock row saved for ${employee.name}`);
    } catch (manualError) {
      setError(manualError.message);
    } finally {
      setBusy(false);
    }
  }

  function startEdit(record) {
    setEditingId(record.id);
    setEditForm({
      action_date: record.action_date,
      action_time: String(record.action_time || "").slice(0, 5),
      employeeKey: record.tbnr,
      direction: record.direction,
    });
  }

  async function saveEdit(record) {
    const employee = employeeByKey.get(editForm.employeeKey) || {
      tbnr: record.tbnr,
      name: record.name,
      type: record.employee_type,
    };
    setBusy(true);
    setError("");
    setMessage("");
    try {
      const payload = await apiJson(`/api/clock/records/${encodeURIComponent(record.id)}`, {
        method: "PATCH",
        body: JSON.stringify({
          action_date: editForm.action_date,
          action_time: editForm.action_time.length === 5 ? `${editForm.action_time}:00` : editForm.action_time,
          tbnr: employee.tbnr,
          name: employee.name,
          employee_type: employee.type || employee.employee_type || "",
          direction: editForm.direction,
        }),
      });
      setEditingId("");
      setRecords(payload.records || []);
      setSessions(payload.sessions || []);
      setSelectedDate(editForm.action_date);
      setMessage("Clock record updated");
    } catch (editError) {
      setError(editError.message);
    } finally {
      setBusy(false);
    }
  }

  async function deleteRecord(record) {
    if (!window.confirm(`Delete ${record.name} ${record.action_time} ${record.direction}?`)) {
      return;
    }
    setBusy(true);
    setError("");
    setMessage("");
    try {
      await apiJson(`/api/clock/records/${encodeURIComponent(record.id)}`, { method: "DELETE" });
      await loadRecords(selectedDate);
      setMessage("Clock record deleted");
    } catch (deleteError) {
      setError(deleteError.message);
    } finally {
      setBusy(false);
    }
  }

  async function refreshDateFromBackup() {
    const date = manual.action_date || selectedDate;
    if (!window.confirm(`Replace app records for ${date} with the spreadsheet backup rows for that date?`)) {
      return;
    }
    setBusy(true);
    setError("");
    setMessage("");
    try {
      const payload = await apiJson("/api/clock/records/import-backup", {
        method: "POST",
        body: JSON.stringify({ date }),
      });
      setSelectedDate(date);
      setRecords(payload.records || []);
      setSessions(payload.sessions || []);
      setMessage(`Backup refreshed for ${date}: ${payload.imported_count || 0} records loaded`);
    } catch (backupError) {
      setError(backupError.message);
    } finally {
      setBusy(false);
    }
  }

  function exportSessionKey(session, index) {
    return session.in_record?.id || session.out_record?.id || `clock-session-${index}`;
  }

  function startExportEdit(session, index) {
    const baseRecord = session.in_record || session.out_record;
    if (!baseRecord) {
      return;
    }
    setExportEditingId(exportSessionKey(session, index));
    setExportEditForm({
      action_date: baseRecord.action_date,
      employeeKey: baseRecord.tbnr,
      in_time: session.in_record ? String(session.in_record.action_time || "").slice(0, 5) : "",
      out_time: session.out_record ? String(session.out_record.action_time || "").slice(0, 5) : "",
    });
  }

  async function saveExportEdit(session) {
    const baseRecord = session.in_record || session.out_record;
    if (!baseRecord) {
      return;
    }
    const employee = employeeByKey.get(exportEditForm.employeeKey) || {
      tbnr: baseRecord.tbnr,
      name: baseRecord.name,
      type: baseRecord.employee_type,
    };
    if (!employee.tbnr || !employee.name) {
      setError("Choose a valid employee");
      return;
    }
    if (!exportEditForm.action_date) {
      setError("Choose a valid date");
      return;
    }
    if (session.in_record && !exportEditForm.in_time) {
      setError("IN time is required");
      return;
    }
    if (session.out_record && !exportEditForm.out_time) {
      setError("OUT time is required");
      return;
    }
    if (!session.in_record && !session.out_record && !exportEditForm.in_time && !exportEditForm.out_time) {
      setError("Add an IN time, an OUT time, or both");
      return;
    }

    setBusy(true);
    setError("");
    setMessage("");
    try {
      if (session.in_record) {
        await apiJson(`/api/clock/records/${encodeURIComponent(session.in_record.id)}`, {
          method: "PATCH",
          body: JSON.stringify({
            action_date: exportEditForm.action_date,
            action_time: exportEditForm.in_time.length === 5 ? `${exportEditForm.in_time}:00` : exportEditForm.in_time,
            tbnr: employee.tbnr,
            name: employee.name,
            employee_type: employee.type || employee.employee_type || "",
            direction: "IN",
          }),
        });
      } else if (exportEditForm.in_time) {
        await apiJson("/api/clock/records", {
          method: "POST",
          body: JSON.stringify({
            employee,
            action_date: exportEditForm.action_date,
            action_time: exportEditForm.in_time.length === 5 ? `${exportEditForm.in_time}:00` : exportEditForm.in_time,
            direction: "IN",
          }),
        });
      }

      if (session.out_record) {
        await apiJson(`/api/clock/records/${encodeURIComponent(session.out_record.id)}`, {
          method: "PATCH",
          body: JSON.stringify({
            action_date: exportEditForm.action_date,
            action_time: exportEditForm.out_time.length === 5 ? `${exportEditForm.out_time}:00` : exportEditForm.out_time,
            tbnr: employee.tbnr,
            name: employee.name,
            employee_type: employee.type || employee.employee_type || "",
            direction: "OUT",
          }),
        });
      } else if (exportEditForm.out_time) {
        await apiJson("/api/clock/records", {
          method: "POST",
          body: JSON.stringify({
            employee,
            action_date: exportEditForm.action_date,
            action_time: exportEditForm.out_time.length === 5 ? `${exportEditForm.out_time}:00` : exportEditForm.out_time,
            direction: "OUT",
          }),
        });
      }

      setExportEditingId("");
      setSelectedDate(exportEditForm.action_date);
      await loadRecords(exportEditForm.action_date);
      setMessage("Clock row updated");
    } catch (editError) {
      setError(editError.message);
    } finally {
      setBusy(false);
    }
  }

  async function deleteExportSession(session) {
    const baseRecord = session.in_record || session.out_record;
    if (!baseRecord) {
      return;
    }
    const label = `${baseRecord.name} ${baseRecord.action_date}`;
    if (!window.confirm(`Delete this clock row for ${label}? This removes the visible IN and OUT actions.`)) {
      return;
    }

    const recordsToDelete = [session.out_record, session.in_record].filter(Boolean);
    setBusy(true);
    setError("");
    setMessage("");
    try {
      for (const record of recordsToDelete) {
        await apiJson(`/api/clock/records/${encodeURIComponent(record.id)}`, { method: "DELETE" });
      }
      setExportEditingId("");
      await loadRecords(selectedDate);
      setMessage("Clock row deleted");
    } catch (deleteError) {
      setError(deleteError.message);
    } finally {
      setBusy(false);
    }
  }

  useEffect(() => {
    if (activeTab !== "clock") {
      return undefined;
    }

    focusScanInput();

    const handlePointerDown = (event) => {
      const input = scanInputRef.current;
      if (!input) {
        return;
      }
      if (event.target instanceof Node && input.contains(event.target)) {
        return;
      }
      focusScanInput();
    };

    const handleKeyDown = (event) => {
      if (busy || event.defaultPrevented || event.metaKey || event.ctrlKey || event.altKey) {
        return;
      }
      const input = scanInputRef.current;
      if (!input || document.activeElement === input) {
        return;
      }

      if (event.key === "Enter") {
        if (scanCodeRef.current) {
          event.preventDefault();
          processScan(scanCodeRef.current);
        }
        return;
      }

      if (event.key === "Backspace") {
        event.preventDefault();
        focusScanInput();
        setScanCode((current) => current.slice(0, -1));
        return;
      }

      if (event.key.length === 1) {
        event.preventDefault();
        focusScanInput();
        setScanCode((current) => `${current}${event.key.toUpperCase()}`);
      }
    };

    document.addEventListener("pointerdown", handlePointerDown, true);
    window.addEventListener("keydown", handleKeyDown, true);
    return () => {
      document.removeEventListener("pointerdown", handlePointerDown, true);
      window.removeEventListener("keydown", handleKeyDown, true);
    };
  }, [activeTab, busy]);

  const exportUrl = `/api/clock/records/export?from=${encodeURIComponent(exportFrom)}&to=${encodeURIComponent(exportTo)}`;

  return (
    <section className="overview-stack clock-page">
      {canManage && (
        <div className="tab-strip clock-tabs">
          <button type="button" className={activeTab === "clock" ? "active" : ""} onClick={() => setActiveTab("clock")}>Clock</button>
          <button type="button" className={activeTab === "extra" ? "active" : ""} onClick={() => setActiveTab("extra")}>Extra</button>
        </div>
      )}

      {error && <div className="notice danger">{error}</div>}
      {message && <div className="notice clock-result">{message}</div>}

      {activeTab === "clock" && (
        <div className="clock-hero">
          <div className="clock-face">
            <strong>{nowLabel}</strong>
          </div>
          <form className="clock-scan-surface" onSubmit={submitScan}>
            <label>
              <span>Badge code</span>
              <input
                ref={scanInputRef}
                value={scanCode}
                onChange={(event) => setScanCode(event.target.value.toUpperCase())}
                onBlur={() => {
                  if (activeTab === "clock") {
                    focusScanInput();
                  }
                }}
                onKeyDown={(event) => {
                  if (event.key === "Enter") {
                    submitScan(event);
                  }
                }}
                placeholder="Scan badge"
                autoFocus
              />
            </label>
          </form>
        </div>
      )}

      {activeTab === "extra" && canManage && (
        <div className="data-table-card">
          <div className="section-header">
            <h2>Extra</h2>
            <button type="button" onClick={() => loadRecords(selectedDate)} disabled={loading}>{loading ? "Loading..." : "Refresh"}</button>
          </div>
          <div className="clock-export-controls">
            <label>
              <span>From date</span>
              <input type="date" value={exportFrom} onChange={(event) => setExportFrom(event.target.value)} />
            </label>
            <label>
              <span>To date</span>
              <input type="date" value={exportTo} onChange={(event) => setExportTo(event.target.value)} />
            </label>
            <a className="button-link" href={exportUrl}>Export range</a>
            <label>
              <span>Show day</span>
              <input
                type="date"
                value={selectedDate}
                onChange={(event) => {
                  setSelectedDate(event.target.value);
                  setManual((current) => ({ ...current, action_date: event.target.value }));
                }}
              />
            </label>
          </div>
          <div className="clock-manual-overview">
            <div className="section-header">
              <h2>Selected date overview</h2>
              <div className="clock-overview-actions">
                <strong>{selectedDateTotalWorked}</strong>
                <button type="button" onClick={refreshDateFromBackup} disabled={busy}>Refresh from backup</button>
              </div>
            </div>
            {employeeLoading && <div className="notice">Loading employees...</div>}
            {loading && <div className="notice">Loading clock records...</div>}
            <div className="table-wrap">
              <table className="data-table">
                <thead>
                  <tr>
                    <th>Date</th>
                    <th>TBNR</th>
                    <th>Name</th>
                    <th>Type</th>
                    <th>IN</th>
                    <th>OUT</th>
                    <th>Worked</th>
                    <th>Source</th>
                    <th>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <td>
                      <input
                        type="date"
                        value={manual.action_date}
                        onChange={(event) => {
                          setManual({ ...manual, action_date: event.target.value });
                          setSelectedDate(event.target.value);
                        }}
                      />
                    </td>
                    <td>
                      <input
                        list="clock-employees"
                        value={manual.employeeKey}
                        onChange={(event) => setManual({ ...manual, employeeKey: event.target.value.toUpperCase() })}
                        placeholder="Badge or name"
                      />
                    </td>
                    <td>{selectedManualEmployee?.name || "-"}</td>
                    <td>{selectedManualEmployee?.type || "-"}</td>
                    <td>
                      <input type="time" value={manual.in_time || ""} onChange={(event) => setManual({ ...manual, in_time: event.target.value })} />
                    </td>
                    <td>
                      <input type="time" value={manual.out_time || ""} onChange={(event) => setManual({ ...manual, out_time: event.target.value })} />
                    </td>
                    <td>-</td>
                    <td>manual</td>
                    <td className="row-actions">
                      <button type="button" onClick={submitManual} disabled={busy}>Add</button>
                      <button
                        type="button"
                        onClick={() => setManual({ employeeKey: "", action_date: selectedDate, in_time: "", out_time: "" })}
                        disabled={busy}
                      >
                        Clear
                      </button>
                    </td>
                  </tr>
                  {sessions.map((session, index) => {
                    const row = session.row || [];
                    const sessionKey = exportSessionKey(session, index);
                    const isEditing = exportEditingId === sessionKey;
                    const selectedEmployee = employeeByKey.get(exportEditForm.employeeKey);
                    return (
                      <tr key={sessionKey}>
                        <td>
                          {isEditing ? (
                            <input type="date" value={exportEditForm.action_date || ""} onChange={(event) => setExportEditForm({ ...exportEditForm, action_date: event.target.value })} />
                          ) : row[0] || "-"}
                        </td>
                        <td>
                          {isEditing ? (
                            <input
                              list="clock-employees"
                              value={exportEditForm.employeeKey || ""}
                              onChange={(event) => setExportEditForm({ ...exportEditForm, employeeKey: event.target.value.toUpperCase() })}
                            />
                          ) : row[1] || "-"}
                        </td>
                        <td>{isEditing ? (selectedEmployee?.name || row[2] || "-") : row[2] || "-"}</td>
                        <td>{isEditing ? (selectedEmployee?.type || row[3] || "-") : row[3] || "-"}</td>
                        <td>
                          {isEditing ? (
                            <input type="time" value={exportEditForm.in_time || ""} onChange={(event) => setExportEditForm({ ...exportEditForm, in_time: event.target.value })} />
                          ) : row[4] || "-"}
                        </td>
                        <td>
                          {isEditing ? (
                            <input type="time" value={exportEditForm.out_time || ""} onChange={(event) => setExportEditForm({ ...exportEditForm, out_time: event.target.value })} />
                          ) : row[5] || "-"}
                        </td>
                        <td>{row[6] || "-"}</td>
                        <td>{row[7] || "-"}</td>
                        <td className="row-actions">
                          {isEditing ? (
                            <>
                              <button type="button" onClick={() => saveExportEdit(session)} disabled={busy}>Save</button>
                              <button type="button" onClick={() => setExportEditingId("")} disabled={busy}>Cancel</button>
                            </>
                          ) : (
                            <>
                              <button type="button" onClick={() => startExportEdit(session, index)} disabled={busy}>Edit</button>
                              <button type="button" onClick={() => deleteExportSession(session)} disabled={busy}>Delete</button>
                            </>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                  {!sessions.length && !loading && (
                    <tr>
                      <td colSpan="9">No clock records for this date.</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
            <datalist id="clock-employees">
              {employees.map((employee) => (
                <option key={employee.tbnr} value={employeeOptionLabel(employee)} />
              ))}
            </datalist>
          </div>
        </div>
      )}
    </section>
  );
}

function SettingsPage({ currentUser }) {
  const [form, setForm] = useState(null);
  const [error, setError] = useState("");
  const [message, setMessage] = useState("");
  const [saving, setSaving] = useState(false);
  const SETTINGS_SECTIONS = [
    { id: "settings-sheets", label: "Sheets" },
    { id: "settings-ukdocs", label: "UKdocs Print" },
    { id: "settings-cmr", label: "CMR / Drive" },
    { id: "settings-mail", label: "Mail" },
  ];
  const { loading: metaLoading, data: metaData, error: metaError } = useFustMeta(Boolean(currentUser));
  const [backups, setBackups] = useState([]);
  const [backupBusy, setBackupBusy] = useState(false);
  const [databaseBusy, setDatabaseBusy] = useState(false);
  const [connectionTest, setConnectionTest] = useState(null);
  const [connectionBusy, setConnectionBusy] = useState(false);
  const [googleBusy, setGoogleBusy] = useState(false);

  async function loadBackups() {
    const payload = await apiJson("/api/fust/backups");
    setBackups(payload.backups || []);
  }

  async function loadConnectionTest() {
    setConnectionBusy(true);
    try {
      const payload = await apiJson("/api/fust/connection-test");
      setConnectionTest(payload);
    } catch (connectionError) {
      setConnectionTest({
        account: { client_email: "", project_id: "" },
        spreadsheet_id: form?.spreadsheet_id || "",
        sheet_name: form?.data_sheet_name || "",
        read_ok: false,
        row_count: 0,
        headers: [],
        error: connectionError.message,
      });
    } finally {
      setConnectionBusy(false);
    }
  }

  useEffect(() => {
    apiJson("/api/fust/settings")
      .then((payload) => setForm({
        ...payload.settings,
        cmr_country_folders_text: cmrFolderLines(payload.settings.cmr_country_folders),
        cmr_manage_usernames_text: (payload.settings.cmr_manage_usernames || []).join("\n"),
      }))
      .catch((settingsError) => setError(settingsError.message));

    loadBackups().catch((backupError) => setError(backupError.message));
    loadConnectionTest().catch(() => {});
  }, []);

  async function submit(event) {
    event.preventDefault();
    setSaving(true);
    setError("");
    setMessage("");
    try {
      const payload = await apiJson("/api/fust/settings", {
        method: "PATCH",
        body: JSON.stringify({
          ...form,
          email_recipients: String(form.email_recipients || "")
            .split(/[\n,;]/)
            .map((value) => value.trim())
            .filter(Boolean),
          cmr_country_folders: parseCmrFolderLines(form.cmr_country_folders_text),
          cmr_manage_usernames: String(form.cmr_manage_usernames_text || "")
            .split(/[\n,;]/)
            .map((value) => value.trim().toLowerCase())
            .filter(Boolean),
        }),
      });
      setForm({
        ...payload.settings,
        email_recipients: payload.settings.email_recipients.join("\n"),
        cmr_country_folders_text: cmrFolderLines(payload.settings.cmr_country_folders),
        cmr_manage_usernames_text: (payload.settings.cmr_manage_usernames || []).join("\n"),
      });
      setMessage("Fust settings saved.");
      await loadConnectionTest();
    } catch (saveError) {
      setError(saveError.message);
    } finally {
      setSaving(false);
    }
  }

  async function createBackup() {
    setBackupBusy(true);
    setError("");
    setMessage("");
    try {
      const payload = await apiJson("/api/fust/backups", {
        method: "POST",
      });
      setBackups(payload.backups || []);
      setMessage(`Backup created: ${payload.backup?.filename || "snapshot saved"}`);
    } catch (backupError) {
      setError(backupError.message);
    } finally {
      setBackupBusy(false);
    }
  }

  async function backfillFustDatabase() {
    if (!window.confirm("Backfill the Fust database now from the current app cache and spreadsheet rows? This will not delete spreadsheet data.")) {
      return;
    }
    setDatabaseBusy(true);
    setError("");
    setMessage("");
    try {
      const payload = await apiJson("/api/fust/database/backfill", {
        method: "POST",
      });
      setMessage(
        `Fust database backfill finished. `
        + `${payload.active_upserted || 0} active action(s) upserted, `
        + `${payload.deleted_marked || 0} deleted action(s) marked. `
        + `Database now has ${payload.database?.active_actions || 0} active and ${payload.database?.deleted_actions || 0} deleted actions.`,
      );
      await loadConnectionTest();
    } catch (backfillError) {
      setError(backfillError.message);
    } finally {
      setDatabaseBusy(false);
    }
  }

  async function restoreMissingFromBackup(filename) {
    if (!window.confirm(`Restore only missing Fust references and CMR/fustbon links from ${filename}? This will not replace the full Fust history.`)) {
      return;
    }
    setBackupBusy(true);
    setError("");
    setMessage("");
    try {
      const payload = await apiJson("/api/fust/backups/restore-missing", {
        method: "POST",
        body: JSON.stringify({ filename }),
      });
      const summary = payload.summary || {};
      setMessage(
        `Backup merge finished from ${payload.filename || filename}. `
        + `${summary.updated || 0} action(s) updated, `
        + `${summary.cmr_restored || 0} CMR, `
        + `${summary.fustbon_restored || 0} fustbon, `
        + `${summary.fustbon_reference_restored || 0} fustbon ref, `
        + `${summary.fustfactuur_reference_restored || 0} fustfactuur ref restored.`,
      );
      await loadBackups();
    } catch (restoreError) {
      setError(restoreError.message);
    } finally {
      setBackupBusy(false);
    }
  }


  async function connectGoogleDrive() {
    setGoogleBusy(true);
    setError("");
    setMessage("");
    try {
      const saved = await apiJson("/api/fust/settings", {
        method: "PATCH",
        body: JSON.stringify({
          ...form,
          email_recipients: String(form.email_recipients || "")
            .split(/[\n,;]/)
            .map((value) => value.trim())
            .filter(Boolean),
          cmr_country_folders: parseCmrFolderLines(form.cmr_country_folders_text),
          cmr_manage_usernames: String(form.cmr_manage_usernames_text || "")
            .split(/[\n,;]/)
            .map((value) => value.trim().toLowerCase())
            .filter(Boolean),
        }),
      });
      setForm({
        ...saved.settings,
        email_recipients: saved.settings.email_recipients.join("\n"),
        cmr_country_folders_text: cmrFolderLines(saved.settings.cmr_country_folders),
        cmr_manage_usernames_text: (saved.settings.cmr_manage_usernames || []).join("\n"),
      });
      const payload = await apiJson("/api/fust/google/auth-url");
      window.location.href = payload.auth_url;
    } catch (googleError) {
      setError(googleError.message);
      setGoogleBusy(false);
    }
  }

  if (!form) {
    return <div className="notice">Loading settings...</div>;
  }

  return (
    <section className="settings-page">
      <form className="settings-form" onSubmit={submit}>
        <div className="tab-strip">
          {SETTINGS_SECTIONS.map((section) => (
            <button
              key={section.id}
              type="button"
              onClick={() => document.getElementById(section.id)?.scrollIntoView({ behavior: "smooth", block: "start" })}
            >
              {section.label}
            </button>
          ))}
        </div>
        <div className="form-grid">
          <div id="settings-sheets" className="wide data-table-card">
            <div className="section-header"><h2>Sheet settings</h2></div>
            <p className="sidebar-note">Main spreadsheet, clock tabs, hal locations, and shared spreadsheet IDs.</p>
          </div>
          <label>
            <span>Spreadsheet ID</span>
            <input
              value={form.spreadsheet_id}
              onChange={(event) => setForm({ ...form, spreadsheet_id: event.target.value })}
            />
          </label>
          <label>
            <span>Data tab</span>
            <input
              value={form.data_sheet_name}
              onChange={(event) => setForm({ ...form, data_sheet_name: event.target.value })}
            />
          </label>
          <label>
            <span>IN tab</span>
            <input
              value={form.in_sheet_name}
              onChange={(event) => setForm({ ...form, in_sheet_name: event.target.value })}
            />
          </label>
          <label>
            <span>OUT tab</span>
            <input
              value={form.out_sheet_name}
              onChange={(event) => setForm({ ...form, out_sheet_name: event.target.value })}
            />
          </label>
          <label>
            <span>Dashboard tab</span>
            <input
              value={form.dashboard_sheet_name}
              onChange={(event) => setForm({ ...form, dashboard_sheet_name: event.target.value })}
            />
          </label>
          <label className="wide">
            <span>Clock spreadsheet ID</span>
            <input
              value={form.clock_spreadsheet_id || ""}
              onChange={(event) => setForm({ ...form, clock_spreadsheet_id: event.target.value })}
              placeholder="Spreadsheet ID for badges and backup"
            />
          </label>
          <label className="wide">
            <span>Hal Locations spreadsheet ID</span>
            <input
              value={form.hal_locations_spreadsheet_id || ""}
              onChange={(event) => setForm({ ...form, hal_locations_spreadsheet_id: event.target.value })}
              placeholder="Defaults to the main spreadsheet ID when left empty"
            />
          </label>
          <label>
            <span>Hal Locations tab</span>
            <input
              value={form.hal_locations_sheet_name || "ERP_PASTE"}
              onChange={(event) => setForm({ ...form, hal_locations_sheet_name: event.target.value })}
              placeholder="ERP_PASTE"
            />
          </label>
          <label>
            <span>Clock employee tab</span>
            <input
              value={form.clock_employee_sheet_name || ""}
              onChange={(event) => setForm({ ...form, clock_employee_sheet_name: event.target.value })}
              placeholder="badges"
            />
          </label>
          <label>
            <span>Clock records tab</span>
            <input
              value={form.clock_records_sheet_name || ""}
              onChange={(event) => setForm({ ...form, clock_records_sheet_name: event.target.value })}
              placeholder="backup"
            />
          </label>
          <div id="settings-ukdocs" className="wide data-table-card">
            <div className="section-header"><h2>UKdocs Print settings</h2></div>
            <p className="sidebar-note">Keep the UKdocs Print source and Gmail pickup together here, so the collection page stays focused on the day workflow.</p>
            <div className="settings-subsection-grid">
              <div className="settings-subsection-card">
                <div className="section-header"><h3>Spreadsheet source</h3></div>
                <div className="form-grid">
                  <label>
                    <span>UKdocs Print spreadsheet ID</span>
                    <input
                      value={form.ukdocs_print_spreadsheet_id || ""}
                      onChange={(event) => setForm({ ...form, ukdocs_print_spreadsheet_id: event.target.value })}
                      placeholder="Spreadsheet ID for PD keuringen sendings"
                    />
                  </label>
                  <label>
                    <span>UKdocs Print tab</span>
                    <input
                      value={form.ukdocs_print_sheet_name || "PD keuringen"}
                      onChange={(event) => setForm({ ...form, ukdocs_print_sheet_name: event.target.value })}
                      placeholder="PD keuringen"
                    />
                  </label>
                </div>
              </div>
              <div className="settings-subsection-card">
                <div className="section-header"><h3>Gmail pickup</h3></div>
                <div className="form-grid">
                  <label className="wide">
                    <span>Connected Gmail account</span>
                    <input value={form.gmail_connected_email || ""} readOnly placeholder="Connect from UKdocs Print page" />
                  </label>
                </div>
                <p className="sidebar-note">The live UKdocs Print page can still reconnect Gmail and run a manual sync whenever needed.</p>
              </div>
            </div>
          </div>
          <div id="settings-cmr" className="wide data-table-card">
            <div className="section-header"><h2>CMR / Drive settings</h2></div>
            <p className="sidebar-note">Google Drive upload, CMR templates, folders, and allowed usernames.</p>
          </div>
          <label className="wide">
            <span>Document country folder IDs</span>
            <textarea
              value={form.cmr_country_folders_text || ""}
              onChange={(event) => setForm({ ...form, cmr_country_folders_text: event.target.value })}
              rows={8}
              placeholder={"FR=google-drive-folder-id\nPT=google-drive-folder-id\nGB=google-drive-folder-id"}
            />
          </label>
          <label className="wide">
            <span>Document fallback folder ID</span>
            <input
              value={form.cmr_fallback_folder_id || ""}
              onChange={(event) => setForm({ ...form, cmr_fallback_folder_id: event.target.value })}
              placeholder="optional fallback folder id"
            />
          </label>
          <label>
            <span>Default CMR template</span>
            <select
              value={form.cmr_default_template_name || ""}
              onChange={(event) => setForm({ ...form, cmr_default_template_name: event.target.value })}
            >
              <option value="">Choose template</option>
              {Array.isArray(form.cmr_available_templates) && form.cmr_available_templates.map((name) => (
                <option key={name} value={name}>{name}</option>
              ))}
            </select>
          </label>
          <label className="wide">
            <span>CMR manager usernames</span>
            <textarea
              value={form.cmr_manage_usernames_text || ""}
              onChange={(event) => setForm({ ...form, cmr_manage_usernames_text: event.target.value })}
              rows={4}
              placeholder={"username1\nusername2"}
            />
          </label>
          <label className="wide">
            <span>CMR data folder</span>
            <input value={form.cmr_data_dir || ""} readOnly />
          </label>
          <label>
            <span>Google OAuth client ID</span>
            <input
              value={form.cmr_google_client_id || ""}
              onChange={(event) => setForm({ ...form, cmr_google_client_id: event.target.value })}
              placeholder="client id"
            />
          </label>
          <label>
            <span>Google OAuth client secret</span>
            <input
              type="password"
              value={form.cmr_google_client_secret || ""}
              onChange={(event) => setForm({ ...form, cmr_google_client_secret: event.target.value })}
              placeholder="client secret"
            />
          </label>
          <div className="wide oauth-connect-row">
            <span>Google Drive upload account: {form.cmr_google_connected_email || "not connected"}</span>
            <button type="button" onClick={connectGoogleDrive} disabled={googleBusy}>
              {googleBusy ? "Connecting..." : "Connect Google Drive"}
            </button>
          </div>
          <div id="settings-mail" className="wide data-table-card">
            <div className="section-header"><h2>Mail settings</h2></div>
            <p className="sidebar-note">Recipients and SMTP account used when papers are ready to send.</p>
          </div>
          <label className="wide">
            <span>Target email recipients</span>
            <textarea
              value={Array.isArray(form.email_recipients) ? form.email_recipients.join("\n") : form.email_recipients}
              onChange={(event) => setForm({ ...form, email_recipients: event.target.value })}
              rows={6}
              placeholder={"name@example.com\nother@example.com"}
            />
          </label>
          <label>
            <span>SMTP host</span>
            <input
              value={form.smtp_host || ""}
              onChange={(event) => setForm({ ...form, smtp_host: event.target.value })}
              placeholder="smtp.gmail.com"
            />
          </label>
          <label>
            <span>SMTP port</span>
            <input
              value={form.smtp_port ?? 587}
              onChange={(event) => setForm({ ...form, smtp_port: event.target.value })}
              placeholder="587"
            />
          </label>
          <label>
            <span>SMTP username</span>
            <input
              value={form.smtp_username || ""}
              onChange={(event) => setForm({ ...form, smtp_username: event.target.value })}
              placeholder="ftereso@gmail.com"
            />
          </label>
          <label>
            <span>SMTP sender email</span>
            <input
              value={form.smtp_from || ""}
              onChange={(event) => setForm({ ...form, smtp_from: event.target.value })}
              placeholder="ftereso@gmail.com"
            />
          </label>
          <label className="wide">
            <span>SMTP app password</span>
            <input
              type="password"
              value={form.smtp_password || ""}
              onChange={(event) => setForm({ ...form, smtp_password: event.target.value })}
              placeholder="16-digit app password"
            />
          </label>
          <label className="checkbox-row">
            <input
              type="checkbox"
              checked={form.smtp_starttls !== false}
              onChange={(event) => setForm({ ...form, smtp_starttls: event.target.checked })}
            />
            <span>Use STARTTLS</span>
          </label>
        </div>
        {message && <div className="notice">{message}</div>}
        {error && <div className="notice danger">{error}</div>}
        <button className="primary" type="submit" disabled={saving}>
          {saving ? "Saving..." : "Save settings"}
        </button>
      </form>

      <InfoPanel
        title="Current live behavior"
        lines={[
          `User storage stays on the Render disk once created by ${currentUser.username}.`,
          "Secret Files are best kept as a seed or restore snapshot.",
          "Fust actions save locally first, then try Sheets and email.",
          "If one sync step fails, the saved action still stays in Render storage.",
          "Target email recipients are where notifications go.",
          "SMTP sender settings are how the app sends the email.",
        ]}
      />

      <div className="data-table-card">
        <div className="section-header">
          <h2>Spreadsheet connection test</h2>
          <div className="row-actions">
            <button type="button" onClick={loadConnectionTest} disabled={connectionBusy || databaseBusy}>
              {connectionBusy ? "Testing..." : "Test again"}
            </button>
            <button type="button" className="primary" onClick={backfillFustDatabase} disabled={databaseBusy || connectionBusy || !connectionTest?.database?.ready}>
              {databaseBusy ? "Backfilling..." : "Backfill Fust to database"}
            </button>
          </div>
        </div>
        {connectionTest && (
          <>
            <p className="sidebar-note">
              Runtime client email: {connectionTest.account?.client_email || "unknown"}
            </p>
            <p className="sidebar-note">
              Project: {connectionTest.account?.project_id || "unknown"}
            </p>
            <p className="sidebar-note">
              Spreadsheet ID: {connectionTest.spreadsheet_id || "(empty)"} | Tab: {connectionTest.sheet_name || "(empty)"}
            </p>
            <p className="sidebar-note">
              Database: {connectionTest.database?.enabled ? (connectionTest.database?.ready ? "connected" : `not ready (${connectionTest.database?.error || "unknown error"})`) : "not configured"}
            </p>
            {!!connectionTest.database_stats && (
              <p className="sidebar-note">
                Database rows: {connectionTest.database_stats.active_actions || 0} active, {connectionTest.database_stats.deleted_actions || 0} deleted, {connectionTest.database_stats.document_rows || 0} document rows
              </p>
            )}
            <p className="sidebar-note">
              Result: {connectionTest.read_ok ? `read ok (${connectionTest.row_count} rows)` : "read failed"}
            </p>
            <p className="sidebar-note">
              Headers: {(connectionTest.headers || []).join(", ") || "none"}
            </p>
            {connectionTest.error && <div className="notice danger">{connectionTest.error}</div>}
          </>
        )}
      </div>

      <div className="data-table-card">
        <h2>Data tab debug</h2>
        {metaLoading && <div className="notice">Loading Data tab preview...</div>}
        {metaError && <div className="notice danger">Unable to read Data tab: {metaError}</div>}
        {!metaLoading && !metaError && (
          <>
            <p className="sidebar-note">
              Source: {metaData?.source || "unknown"} | Rows read: {metaData?.raw_row_count || 0} | Usable records: {metaData?.records?.length || 0}
            </p>
            <p className="sidebar-note">
              Headers: {(metaData?.headers || []).join(", ") || "none detected"}
            </p>
            <div className="table-wrap">
              <table className="data-table">
                <thead>
                  <tr>
                    <th>klantnaam</th>
                    <th>Country</th>
                    <th>klantcode connect</th>
                  </tr>
                </thead>
                <tbody>
                  {(metaData?.sample_records || []).map((record, index) => (
                    <tr key={`${record.customer_name}-${record.country}-${record.customer_code}-${index}`}>
                      <td>{record.customer_name}</td>
                      <td>{record.country}</td>
                      <td>{record.customer_code || record.connect_name}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </>
        )}
      </div>

      <div className="data-table-card">
        <div className="section-header">
          <h2>Fust backups</h2>
          <button type="button" className="primary" onClick={createBackup} disabled={backupBusy}>
            {backupBusy ? "Creating..." : "Create backup"}
          </button>
        </div>
        <p className="sidebar-note">
          Snapshots are stored on the Render persistent disk and include current Fust settings plus all saved actions.
        </p>
        <div className="table-wrap">
          <table className="data-table">
            <thead>
              <tr>
                <th>Filename</th>
                <th>Created</th>
                <th>Size</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {backups.map((backup) => (
                <tr key={backup.filename}>
                  <td>{backup.filename}</td>
                  <td>{formatTimestamp(backup.created_at)}</td>
                  <td>{Math.round((backup.size_bytes || 0) / 1024)} KB</td>
                  <td>
                    <a href={backup.download_path}>Download</a>
                    {" "}
                    <button type="button" onClick={() => restoreMissingFromBackup(backup.filename)} disabled={backupBusy}>
                      Restore missing info
                    </button>
                  </td>
                </tr>
              ))}
              {!backups.length && (
                <tr>
                  <td colSpan="4">No Fust backups created yet.</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </section>
  );
}

function InfoPanel({ title, lines }) {
  return (
    <article className="info-panel">
      <h2>{title}</h2>
      <div>
        {lines.map((line) => (
          <p key={line}>{line}</p>
        ))}
      </div>
    </article>
  );
}

function SyncPanel({ status }) {
  if (!status?.state) {
    return <p className="sync idle">Background sync idle</p>;
  }

  if (status.state === "running") {
    return <p className="sync running">Running {String(status.mode || "").replaceAll("_", " ")} since {formatTimestamp(status.started_at)}</p>;
  }

  if (status.state === "completed") {
    return <p className="sync done">Background sync completed at {formatTimestamp(status.updated_at)}</p>;
  }

  if (status.state === "failed") {
    return <p className="sync failed">Background sync failed: {status.error || "unknown error"}</p>;
  }

  return <p className="sync idle">Background sync {status.state}</p>;
}

function ParseErrors({ errors }) {
  if (!errors.length) {
    return <p className="sidebar-note">No skipped folders.</p>;
  }

  return (
    <details className="parse-errors">
      <summary>{errors.length} skipped folders</summary>
      <div>
        {errors.slice(0, 80).map((error) => (
          <p key={`${error.folder_id}-${error.reason}`}>
            <strong>{error.folder_name}</strong>
            <span>{error.carrier}</span>
          </p>
        ))}
      </div>
    </details>
  );
}

function Metric({ label, value }) {
  return (
    <div className="metric">
      <span>{label}</span>
      <strong>{value}</strong>
    </div>
  );
}

function CustomerGroup({ group, expanded, onToggle, onOpenPhoto }) {
  const totalImages = group.runs.reduce((total, run) => total + (run.images?.length || 0), 0);
  const carriers = [...new Set(group.runs.map((run) => run.carrier).filter(Boolean))].sort();

  return (
    <article className={`customer ${expanded ? "open" : ""}`}>
      <button className="customer-header" onClick={onToggle} aria-expanded={expanded}>
        <span>
          <strong>{group.customer_code}</strong>
          <small>
            {group.runs.length} runs | {totalImages} images | Carriers: {carriers.join(", ")}
          </small>
        </span>
        <b>{expanded ? "Collapse" : "Expand"}</b>
      </button>

      {expanded && (
        <div className="run-list">
          {group.runs.map((run) => (
            <RunCard key={run.folder_id} run={run} onOpenPhoto={onOpenPhoto} />
          ))}
        </div>
      )}
    </article>
  );
}

function RunCard({ run, onOpenPhoto }) {
  const hasQrInfo = run.qr_info && run.qr_info !== "No QR info found";

  return (
    <section className="run-card">
      <header>
        <h2>{run.carrier} | {run.run_id || "No run ID"}</h2>
        <p><strong>Carrier:</strong> {run.carrier}</p>
        <p><strong>Run ID:</strong> {run.run_id || "N/A"}</p>
        <p><strong>Folder:</strong> <code>{run.folder_name}</code></p>
      </header>

      {hasQrInfo && (
        <div className="qr">
          <strong>QR Info</strong>
          {run.qr_source === "filename" ? <code>{run.qr_info}</code> : <pre>{run.qr_info}</pre>}
        </div>
      )}

      {run.images?.length ? (
        <div className="gallery">
          {run.images.map((image, index) => (
            <button
              className="photo-card"
              key={image.id}
              onClick={() => onOpenPhoto(run, index)}
              aria-label={`Open ${image.name}`}
            >
              <PhotoImage image={image} run={run} alt={image.name} loading="lazy" />
              <span>{image.name}</span>
            </button>
          ))}
        </div>
      ) : (
        <p className="empty">No image files found in this run folder.</p>
      )}
    </section>
  );
}

function Lightbox({ photos, index, onChange, onClose }) {
  const active = photos[index];

  useEffect(() => {
    function handleKeydown(event) {
      if (event.key === "ArrowLeft") {
        event.preventDefault();
        onChange((index - 1 + photos.length) % photos.length);
      }
      if (event.key === "ArrowRight") {
        event.preventDefault();
        onChange((index + 1) % photos.length);
      }
      if (event.key === "Escape") {
        event.preventDefault();
        onClose();
      }
    }

    window.addEventListener("keydown", handleKeydown);
    return () => window.removeEventListener("keydown", handleKeydown);
  }, [index, onChange, onClose, photos.length]);

  function previous() {
    onChange((index - 1 + photos.length) % photos.length);
  }

  function next() {
    onChange((index + 1) % photos.length);
  }

  return (
    <div className="lightbox" role="dialog" aria-modal="true" onClick={onClose}>
      <button className="lightbox-close" onClick={onClose} aria-label="Close photo viewer">x</button>
      <button className="lightbox-nav previous" onClick={(event) => { event.stopPropagation(); previous(); }} aria-label="Previous photo">
        ‹
      </button>
      <figure onClick={(event) => event.stopPropagation()}>
        <PhotoImage image={active.image} run={active.run} alt={active.image.name} />
        <figcaption>
          {index + 1} / {photos.length} - {active.image.name}
        </figcaption>
      </figure>
      <button className="lightbox-nav next" onClick={(event) => { event.stopPropagation(); next(); }} aria-label="Next photo">
        ›
      </button>
    </div>
  );
}

createRoot(document.getElementById("root")).render(<App />);
