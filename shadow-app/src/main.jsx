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
  CMR_VIEW: "cmr:view",
  HAL_LOCATIONS_VIEW: "hal_locations:view",
  EXPEDITION_STICKERS_VIEW: "expedition_stickers:view",
  CMR_MANAGE: "cmr:manage",
  CLOCK_VIEW: "clock:view",
  CLOCK_MANAGE: "clock:manage",
  USERS_MANAGE: "users:manage",
  SETTINGS_MANAGE: "settings:manage",
  UKDOCS_VIEW: "ukdocs:view",
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
  { key: "ukdocsprint", label: "UKdocs Print", permission: PERMISSIONS.UKDOCS_VIEW },
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
        caption: "Capture IN and OUT movements, then review balances and recent actions.",
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
        caption: "Save shared planning and split files, refresh ERP_PASTE, and generate expedition sticker PDFs for the team.",
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
    return document.error || "Upload failed";
  }
  return "Missing";
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
      setMessage("Shared expedition source files saved.");
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
    } catch (sheetError) {
      setHalSessionId("");
      setHalSummary(null);
      setError(sheetError.message);
      setMessage("");
    } finally {
      setSheetBusy(false);
    }
  }

  async function generateExpeditionStickers() {
    if (!halSessionId) {
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
        body: JSON.stringify({ id: halSessionId }),
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

  return (
    <section className="overview-stack">
      <article className="panel hal-panel">
        <div className="section-header">
          <div>
            <h2>Shared source files</h2>
            <p>The planner uploads the latest planning and split files here once. Everyone else can reuse them later without uploading again.</p>
          </div>
          <button type="button" onClick={loadState} disabled={loadingState}>
            {loadingState ? "Refreshing..." : "Refresh saved state"}
          </button>
        </div>
        <div className="form-grid">
          <label>
            <span>Planning file</span>
            <input type="file" accept=".xlsx,.xls,.csv" onChange={(event) => setPlanningFile(event.target.files?.[0] || null)} />
          </label>
          <label>
            <span>Split file</span>
            <input type="file" accept=".xlsx,.xls,.csv" onChange={(event) => setSplitFile(event.target.files?.[0] || null)} />
          </label>
        </div>
        <div className="row-actions hal-actions-row">
          <button type="button" className="primary" onClick={saveSources} disabled={savingSources}>
            {savingSources ? "Saving..." : "Save shared files"}
          </button>
        </div>
        <div className="table-wrap">
          <table className="data-table">
            <thead>
              <tr>
                <th>Source</th>
                <th>Saved file</th>
                <th>Rows</th>
                <th>Saved by</th>
                <th>Saved at</th>
              </tr>
            </thead>
            <tbody>
              {[
                ["Planning", savedState?.planning_file, savedState?.planning_summary],
                ["Split", savedState?.split_file, savedState?.split_summary],
              ].map(([label, file, summary]) => (
                <tr key={label}>
                  <td>{label}</td>
                  <td>{file?.original_name || "-"}</td>
                  <td>{summary?.row_count || 0}</td>
                  <td>{file?.saved_by || "-"}</td>
                  <td>{file?.saved_at ? formatTimestamp(file.saved_at) : "-"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </article>

      <article className="panel hal-panel">
        <div className="section-header">
          <div>
            <h2>Live halindeling</h2>
            <p>Reuse the same ERP_PASTE spreadsheet source as Hal Locations so the latest location lookup is always loaded right before generation.</p>
          </div>
        </div>
        <div className="row-actions hal-actions-row">
          <button type="button" className="primary" onClick={loadHalindelingFromSheet} disabled={sheetBusy}>
            {sheetBusy ? "Loading sheet..." : "Load ERP_PASTE from spreadsheet"}
          </button>
        </div>
        {savedState?.sheet_source?.spreadsheet_id ? (
          <p className="sidebar-note">
            Source: {savedState.sheet_source.sheet_name || "ERP_PASTE"} ({savedState.sheet_source.spreadsheet_id})
          </p>
        ) : null}
        {halSummary ? (
          <p className="sidebar-note">
            Loaded {halSummary.totalRows || 0} rows, {(halSummary.locPrefixes || []).length} location prefixes, and {(halSummary.custPrefixes || []).length} customer prefixes.
          </p>
        ) : null}
      </article>

      <article className="panel hal-panel">
        <div className="section-header">
          <div>
            <h2>Generate expedition stickers</h2>
            <p>The PDFs follow the old sticker app flow: one file per split / truck plus one overig file when rows have no split.</p>
          </div>
        </div>
        <div className="row-actions hal-actions-row">
          <button type="button" className="primary" onClick={generateExpeditionStickers} disabled={generating || !halSessionId}>
            {generating ? "Generating..." : "Generate PDFs"}
          </button>
          <button type="button" disabled={!generatedFiles.length} onClick={() => generatedFiles.forEach((file) => downloadBase64File(file.name, file.content_base64, file.mime_type))}>
            Download all
          </button>
        </div>
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

function CmrBatchPrintView({ customers, buildPrintPage, defaultNatureOfGoods }) {
  const [search, setSearch] = useState("");
  const [selectedNames, setSelectedNames] = useState([]);
  const [overrides, setOverrides] = useState({});
  const visibleCustomers = useMemo(() => sortByName(customers).filter((item) => `${item.name} ${item.address || ""}`.toLowerCase().includes(search.trim().toLowerCase())), [customers, search]);

  function toggle(name) {
    setSelectedNames((current) => current.includes(name) ? current.filter((item) => item !== name) : [...current, name]);
  }

  function openBatch(autoPrint) {
    const pages = selectedNames
      .map((name) => visibleCustomers.find((item) => item.name === name) || customers.find((item) => item.name === name))
      .filter(Boolean)
      .map((customer) => buildPrintPage(customer, { natureOfGoods: overrides[customer.name] || defaultNatureOfGoods }));
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
      </div>
      <div className="cmr-batch-list-table">
        {visibleCustomers.map((customer) => (
          <div key={customer.name} className="cmr-batch-row">
            <div>
              <strong>{customer.name}</strong>
              <div className="sidebar-note">{customer.address || "-"}</div>
            </div>
            <label>
              <span>Field 9</span>
              <textarea rows={4} value={overrides[customer.name] || defaultNatureOfGoods} onChange={(event) => setOverrides((current) => ({ ...current, [customer.name]: event.target.value }))} />
            </label>
            <label className="checkbox-row">
              <input type="checkbox" checked={selectedNames.includes(customer.name)} onChange={() => toggle(customer.name)} />
              <span>Add</span>
            </label>
          </div>
        ))}
      </div>
      <div className="row-actions spread-actions">
        <button type="button" onClick={() => openBatch(false)}>Preview selected</button>
        <button type="button" className="primary" onClick={() => openBatch(true)}>Print selected</button>
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
  const visibleMenus = canManage ? CMR_MENU_DEFINITIONS : CMR_MENU_DEFINITIONS.filter((item) => item.key === "cmrprint");
  const settings = cmrData?.settings || {};

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

  const selectedTemplateName = settings.cmr_default_template_name || templates[0]?.name || "";
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
    setSaving(true);
    setMessage("");
    setSaveError("");
    try {
      const payload = await apiJson("/api/cmrprint/app-data", {
        method: "PATCH",
        body: JSON.stringify({
          customers,
          exporters,
          transport_infos: transportInfos,
          loading_places: loadingPlaces,
        }),
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
    refresh();
  }

  async function deleteTemplate(templateName) {
    const payload = await apiJson(`/api/cmrprint/template/${encodeURIComponent(templateName)}`, { method: "DELETE" });
    setDraftData((current) => ({ ...current, templates: payload.templates || [] }));
    refresh();
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
            <label><span>Exporter</span><input value={exporter?.name || "-"} readOnly /></label>
            <label><span>Transport</span><input value={transportInfo?.name || "-"} readOnly /></label>
            <label className="wide"><span>Customer block</span><textarea rows={4} value={buildCmrCustomerBlock(customer)} readOnly /></label>
          </div>
          <div className="cmr-print-toolbar"><div className="row-actions spread-actions"><button type="button" onClick={() => openCurrentCustomerPrint(false)}>Preview Print</button><button type="button" className="primary" onClick={() => openCurrentCustomerPrint(true)}>Print CMR</button></div></div>
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
      {canManage && activeMenu === "batch" && <CmrBatchPrintView customers={customers} buildPrintPage={buildPrintPageForCustomer} defaultNatureOfGoods={manualValues.natureOfGoods} />}
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
  ["match_hub_code", "Hub code match"],
  ["match_remark", "Remark match"],
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
  ["default_invoice_language_text", "Default invoice language / text", "textarea"],
  ["default_document_references", "Default document references", "textarea"],
];

const UKDOCS_CUSTOMER_INVOICE_VISIBILITY_FIELDS = [
  ["show_invoice_vat_number", "Show VAT number on invoice"],
  ["show_invoice_eori_number", "Show EORI number on invoice"],
  ["show_invoice_importer_number", "Show importer / DAN number on invoice"],
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
    default_invoice_language_text: "",
    default_document_references: "",
    show_invoice_vat_number: true,
    show_invoice_eori_number: true,
    show_invoice_importer_number: true,
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

function findUkdocsCustomerMatch(customers, collection) {
  const hubCode = normalizeUkdocsMatchToken(collection?.hub_code);
  const remark = normalizeUkdocsMatchToken(collection?.remark);
  let bestMatch = null;
  let bestScore = 0;
  for (const customer of customers || []) {
    const customerHubCode = normalizeUkdocsMatchToken(customer?.match_hub_code);
    const customerRemark = normalizeUkdocsMatchToken(customer?.match_remark);
    if (!customerHubCode && !customerRemark) {
      continue;
    }
    let score = 0;
    if (customerHubCode) {
      if (!hubCode || customerHubCode !== hubCode) {
        continue;
      }
      score += 2;
    }
    if (customerRemark) {
      if (!remark || !remark.includes(customerRemark)) {
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
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const [state, setState] = useState(null);
  const [customerDraft, setCustomerDraft] = useState(emptyUkdocsCustomer());
  const [shipmentDraft, setShipmentDraft] = useState(emptyUkdocsShipmentDraft());
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
  const availablePrintCollections = printCollections.filter((item) => item.source === "sheet" || item.reference_connect);
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

  function applyPrintCollection(collectionId) {
    const collection = printCollections.find((item) => item.id === collectionId) || null;
    const matchedCustomer = (collection?.customer_id && customers.find((item) => item.id === collection.customer_id)) || findUkdocsCustomerMatch(customers, collection);
    setShipmentDraft((current) => ({
      ...current,
      print_collection_id: collectionId,
      customer_id: matchedCustomer?.id || current.customer_id,
      reference_connect: collection?.reference_connect || current.reference_connect,
      shipment_date: collection?.shipment_date || current.shipment_date,
      trailer_number: current.trailer_number || collection?.trailer_number || "",
      truck_number: current.truck_number || collection?.truck_number || "",
      delivery_terms: matchedCustomer?.default_delivery_terms || matchedCustomer?.export_defaults?.delivery_terms || current.delivery_terms,
      uk_arrival_port: matchedCustomer?.default_uk_arrival_port || current.uk_arrival_port,
      currency: matchedCustomer?.default_currency || matchedCustomer?.export_defaults?.currency || current.currency,
      owner: matchedCustomer?.default_owner || current.owner,
      importer: matchedCustomer?.default_importer || matchedCustomer?.importer_number || matchedCustomer?.eori_number || matchedCustomer?.export_defaults?.importer_field || current.importer,
      delivery_terms_city: matchedCustomer?.export_defaults?.delivery_terms_city || matchedCustomer?.default_city || current.delivery_terms_city,
      regulation: matchedCustomer?.export_defaults?.regulation || current.regulation,
      destination_country: matchedCustomer?.export_defaults?.destination_country || current.destination_country,
      customs_office_of_exit: matchedCustomer?.export_defaults?.customs_office_of_exit || current.customs_office_of_exit,
      location: matchedCustomer?.export_defaults?.location || current.location,
      border_transport_mode: matchedCustomer?.export_defaults?.border_transport_mode || current.border_transport_mode,
      border_transport_nationality: matchedCustomer?.export_defaults?.border_transport_nationality || current.border_transport_nationality,
      freight_costs: matchedCustomer?.export_defaults?.freight_costs || current.freight_costs,
      insurance: matchedCustomer?.export_defaults?.insurance || current.insurance,
      vessel: matchedCustomer?.export_defaults?.vessel_field || current.vessel,
      notes: collection
        ? [current.notes, `Reference connect: ${collection.reference_connect || "-"}`, collection.city_name ? `Sending city: ${collection.city_name}` : "", collection.hub_code ? `Hub: ${collection.hub_code}` : ""]
          .filter(Boolean)
          .join("\n")
        : current.notes,
    }));
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
      setState((current) => ({ ...current, shipments: payload.shipments }));
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
          <button key={key} type="button" className={activeMenu === key ? "active" : ""} onClick={() => setActiveMenu(key)}>{label}</button>
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
            <label className="wide">
              <span>Available sending</span>
              <select value={shipmentDraft.print_collection_id || ""} onChange={(event) => applyPrintCollection(event.target.value)}>
                <option value="">Choose a sending from UKdocs Print</option>
                {availablePrintCollections.map((collection) => (
                  <option key={collection.id} value={collection.id}>
                    {`${collection.shipment_date || "-"} | ${collection.reference_connect || "-"} | ${collection.city_name || collection.customer_name || "-"} | ${collection.hub_code || "-"}`}
                  </option>
                ))}
              </select>
            </label>
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
            <button type="button" onClick={resetDrafts} disabled={saving}>New blank shipment</button>
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
          <div className="table-wrap"><table className="data-table"><thead><tr><th>Name</th><th>Hub match</th><th>Remark match</th><th>Delivery terms</th><th>UK port</th><th>Currency</th><th>VAT</th><th>Actions</th></tr></thead><tbody>{customers.map((customer) => <tr key={customer.id}><td>{customer.customer_name}</td><td>{customer.match_hub_code || "-"}</td><td>{customer.match_remark || "-"}</td><td>{customer.default_delivery_terms || customer.export_defaults?.delivery_terms || "-"}</td><td>{customer.default_uk_arrival_port || "-"}</td><td>{customer.default_currency || customer.export_defaults?.currency || "-"}</td><td>{customer.vat_number || "-"}</td><td className="row-actions"><button type="button" onClick={() => startEditCustomer(customer)}>Edit</button></td></tr>)}{!customers.length && <tr><td colSpan="8">No UKdocs customers saved yet.</td></tr>}</tbody></table></div>
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

function UkdocsPrintPage({ currentUser }) {
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [state, setState] = useState(null);
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const [selectedCollectionId, setSelectedCollectionId] = useState("");
  const [notesDraft, setNotesDraft] = useState("");
  const [gmailQuery, setGmailQuery] = useState("has:attachment newer_than:30d");
  const [gmailSyncResults, setGmailSyncResults] = useState([]);
  const [gmailBusy, setGmailBusy] = useState(false);
  const [gmailSettings, setGmailSettings] = useState({ gmail_connected_email: "" });
  const [sheetSyncDate, setSheetSyncDate] = useState(new Date().toISOString().slice(0, 10));
  const [sheetBusy, setSheetBusy] = useState(false);
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
  const selectedCollection = collections.find((item) => item.id === selectedCollectionId || item.shipment_id === selectedCollectionId) || collections[0] || null;
  const selectedPhytoFiles = selectedCollection?.documents?.phyto_files || [];

  useEffect(() => {
    if (selectedCollection?.id && selectedCollection.id !== selectedCollectionId) {
      setSelectedCollectionId(selectedCollection.id);
    }
  }, [selectedCollection?.id, selectedCollectionId]);

  useEffect(() => {
    setNotesDraft(selectedCollection?.notes || "");
  }, [selectedCollection?.id, selectedCollection?.notes]);

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
      setState((current) => ({ ...current, print_collections: nextCollections }));
      setSelectedCollectionId(nextCollections[0]?.id || "");
      setNotesDraft(nextCollections[0]?.notes || "");
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

  async function syncGmail() {
    setGmailBusy(true);
    setError("");
    setMessage("");
    try {
      const payload = await apiJson("/api/ukdocs-print/gmail/sync", {
        method: "POST",
        body: JSON.stringify({ query: gmailQuery }),
      });
      setState((current) => ({ ...current, print_collections: payload.print_collections || current?.print_collections || [] }));
      setGmailSyncResults(payload.results || []);
      setMessage(`Gmail sync finished. ${payload.matched || 0} matched, ${payload.unmatched || 0} unmatched, ${payload.skipped || 0} skipped.`);
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
      <div className="notice">Use the spreadsheet to load the day sendings first. Then link UKdocs generated files to one of those sendings, and let Gmail attach the phytosanitary and second export documents onto the same record.</div>

      <div className="data-table-card ukdocs-stack">
        <div className="section-header"><h2>Today Sendings Spreadsheet</h2></div>
        <div className="form-grid">
          <label><span>Date to import</span><input type="date" value={sheetSyncDate} onChange={(event) => setSheetSyncDate(event.target.value)} /></label>
          <label><span>Spreadsheet ID</span><input value={gmailSettings.ukdocs_print_spreadsheet_id || ""} readOnly placeholder="Not set in Settings" /></label>
          <label><span>Spreadsheet tab</span><input value={gmailSettings.ukdocs_print_sheet_name || ""} readOnly placeholder="Not set in Settings" /></label>
        </div>
        <div className="row-actions spread-actions">
          <button type="button" className="primary" onClick={syncSheetSendings} disabled={sheetBusy}>{sheetBusy ? "Loading..." : "Load sendings from spreadsheet"}</button>
        </div>
        <div className="notice">This imports the sending list for the selected day from the PD spreadsheet. Then UKdocs shipments can link to one of these sendings, and Gmail can match the phytosanitary PDF by reference connect.</div>
      </div>

      <div className="data-table-card ukdocs-stack">
        <div className="section-header"><h2>Gmail Inbox Pickup</h2></div>
        <div className="form-grid">
          <label><span>Connected Gmail account</span><input value={gmailSettings.gmail_connected_email || ""} readOnly placeholder="Not connected yet" /></label>
          <label className="wide"><span>Gmail search query</span><input value={gmailQuery} onChange={(event) => setGmailQuery(event.target.value)} placeholder="has:attachment newer_than:30d" /></label>
        </div>
        <div className="row-actions spread-actions">
          {canManageSettings && <button type="button" onClick={connectGmail} disabled={gmailBusy}>{gmailBusy ? "Connecting..." : "Connect Gmail"}</button>}
          <button type="button" className="primary" onClick={syncGmail} disabled={gmailBusy}>{gmailBusy ? "Syncing..." : "Sync Gmail attachments"}</button>
        </div>
      <div className="notice">The sync checks the email body and subject for reference connect first, then invoice numbers, then truck or trailer registration. NVWA / e-CertNL emails are treated as phytosanitary documents automatically. Files only fill empty slots automatically, so manual uploads stay safe.</div>
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
          <div className="table-wrap">
            <table className="data-table">
              <thead><tr><th>Date</th><th>Connect ref</th><th>City / customer</th><th>Invoices</th><th>Truck</th><th>Status</th><th>Actions</th></tr></thead>
              <tbody>
                {collections.map((collection) => {
                  const status = ukdocsPrintStatusDefinition(collection.status);
                  return (
                    <tr key={collection.id}>
                      <td>{collection.shipment_date || "-"}</td>
                      <td>{collection.reference_connect || "-"}</td>
                      <td>{collection.city_name || collection.customer_name || "-"}</td>
                      <td>{collection.invoice_numbers || "-"}</td>
                      <td>{collection.truck_number || collection.trailer_number || "-"}</td>
                      <td><span className={`ukdocs-status-badge ${status.tone}`}>{status.label}</span></td>
                      <td className="row-actions"><button type="button" onClick={() => setSelectedCollectionId(collection.id)}>Open</button><button type="button" onClick={() => deleteCollection(collection.id)}>Delete</button></td>
                    </tr>
                  );
                })}
                {!collections.length && <tr><td colSpan="7">No spreadsheet sendings or UKdocs-linked collections are loaded yet.</td></tr>}
              </tbody>
            </table>
          </div>
        </div>

        <div className="data-table-card ukdocs-stack">
          <div className="section-header">
            <h2>Collection detail</h2>
            {selectedCollection && <div className="row-actions"><div className={`ukdocs-status-badge ${ukdocsPrintStatusDefinition(selectedCollection.status).tone}`}>{ukdocsPrintStatusDefinition(selectedCollection.status).label}</div><button type="button" onClick={() => deleteCollection(selectedCollection.id)}>Delete</button></div>}
          </div>

          {!selectedCollection && <div className="notice">Choose a generated shipment first.</div>}

          {selectedCollection && (
            <>
              <div className="form-grid">
                <label><span>Shipment reference</span><input value={selectedCollection.shipment_reference || ""} readOnly /></label>
                <label><span>Reference connect</span><input value={selectedCollection.reference_connect || ""} readOnly /></label>
                <label><span>Shipment date</span><input value={selectedCollection.shipment_date || ""} readOnly /></label>
                <label><span>City / customer</span><input value={selectedCollection.city_name || selectedCollection.customer_name || ""} readOnly /></label>
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

              <div className="ukdocs-upload-grid">
                {UKDOCS_PRINT_DOCUMENTS.map((documentDefinition) => {
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
                                <a key={`${phytoFile.storage_name}-${index}`} href={`/api/ukdocs-print/collections/${encodeURIComponent(selectedCollection.id)}/documents/phyto/${index}`}>
                                  {phytoFile.original_name || `Phyto ${index + 1}`}
                                </a>
                              ))}
                            </div>
                          )}
                        </>
                      ) : (
                        <>
                          <small>{document?.original_name ? `${document.original_name} saved ${formatTimestamp(document.saved_at)}` : "No file saved yet."}</small>
                          {document?.storage_name && <div className="row-actions"><a href={`/api/ukdocs-print/collections/${encodeURIComponent(selectedCollection.id)}/documents/${documentDefinition.key}`}>Download</a></div>}
                        </>
                      )}
                    </div>
                  );
                })}
              </div>

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
          <p>{heading.caption}</p>
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
          <p className="eyebrow">SnappySjaak</p>
          <h1>Sjaak vd Vijver App</h1>
          <p className="sidebar-note">Photos stays separate while Fust gets its own protected workspace.</p>
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
        {page === "ukdocsprint" && <UkdocsPrintPage currentUser={auth.user} />}
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
  if (tab === "last-actions") {
    return "Last actions";
  }
  return tab.toUpperCase();
}

function FustPage({ currentUser, menuVersion }) {
  const visibleTabs = [
    hasPermission(currentUser, PERMISSIONS.FUST_IN) ? "in" : null,
    hasPermission(currentUser, PERMISSIONS.FUST_OUT) ? "out" : null,
    hasPermission(currentUser, PERMISSIONS.FUST_OVERVIEW) ? "overview" : null,
    hasPermission(currentUser, PERMISSIONS.FUST_OVERVIEW) ? "last-actions" : null,
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
    return <div className="notice">This account can open Fust but does not yet have an IN, OUT, or Overview action assigned.</div>;
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

function FustLastActions({ loading, actions, onRefresh }) {
  const [busyActionId, setBusyActionId] = useState("");
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const [editingActionId, setEditingActionId] = useState("");
  const [editForm, setEditForm] = useState(null);
  const [typeFilter, setTypeFilter] = useState("");
  const [dateFilter, setDateFilter] = useState("");
  const [countryFilter, setCountryFilter] = useState("");
  const [customerFilter, setCustomerFilter] = useState("");

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
      setMessage("Local action deleted.");
      onRefresh();
    } catch (deleteError) {
      setError(deleteError.message);
    } finally {
      setBusyActionId("");
    }
  }

  const typeOptions = [...new Set(actions.map((action) => action.type).filter(Boolean))]
    .sort((left, right) => left.localeCompare(right));
  const countryOptions = [...new Set(actions.map((action) => action.country).filter(Boolean))]
    .sort((left, right) => left.localeCompare(right));
  const customerOptions = [...new Set(actions.map((action) => action.customer_name).filter(Boolean))]
    .sort((left, right) => left.localeCompare(right));
  const visibleActions = actions.filter((action) => {
    if (typeFilter && action.type !== typeFilter) {
      return false;
    }
    if (dateFilter && String(action.action_date || "") !== dateFilter) {
      return false;
    }
    if (countryFilter && action.country !== countryFilter) {
      return false;
    }
    if (customerFilter && action.customer_name !== customerFilter) {
      return false;
    }
    return true;
  });

  if (loading) {
    return <div className="notice">Loading last actions...</div>;
  }

  return (
    <div className="overview-stack">
      {message && <div className="notice">{message}</div>}
      {error && <div className="notice danger">{error}</div>}

      <div className="data-table-card">
        <div className="section-header">
          <h2>Last actions</h2>
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
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {visibleActions.map((action) => {
                const isEditing = editingActionId === action.id && editForm;
                const canModify = action.created_by !== "spreadsheet";
                return (
                  <tr key={action.id}>
                    <td>
                      {isEditing ? (
                        <select value={editForm.type} onChange={(event) => setEditForm({ ...editForm, type: event.target.value })}>
                          <option value="IN">IN</option>
                          <option value="OUT">OUT</option>
                        </select>
                      ) : action.type}
                    </td>
                    <td>
                      {isEditing ? (
                        <input type="date" value={editForm.action_date} onChange={(event) => setEditForm({ ...editForm, action_date: event.target.value })} />
                      ) : action.action_date}
                    </td>
                    <td>
                      {isEditing ? (
                        <input value={editForm.country} onChange={(event) => setEditForm({ ...editForm, country: event.target.value })} />
                      ) : action.country}
                    </td>
                    <td>
                      {isEditing ? (
                        <input value={editForm.customer_name} onChange={(event) => setEditForm({ ...editForm, customer_name: event.target.value })} />
                      ) : action.customer_name}
                    </td>
                    <td>
                      {isEditing ? (
                        <input value={editForm.connect_name} onChange={(event) => setEditForm({ ...editForm, connect_name: event.target.value })} />
                      ) : action.connect_name}
                    </td>
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
                              metrics: {
                                ...editForm.metrics,
                                [metric]: Number(event.target.value || 0),
                              },
                            })}
                          />
                        ) : (action.metrics?.[metric] || 0)}
                      </td>
                    ))}
                    <td><DocumentStatus action={action} /></td>
                    <td>
                      {isEditing ? (
                        <input value={editForm.remark} onChange={(event) => setEditForm({ ...editForm, remark: event.target.value })} />
                      ) : (action.remark || "-")}
                    </td>
                    <td>
                      {isEditing ? (
                        <input value={editForm.fustbon_reference} onChange={(event) => setEditForm({ ...editForm, fustbon_reference: event.target.value })} />
                      ) : (action.fustbon_reference || "-")}
                    </td>
                    <td>
                      {isEditing ? (
                        <input value={editForm.fustfactuur_reference} onChange={(event) => setEditForm({ ...editForm, fustfactuur_reference: event.target.value })} />
                      ) : (action.fustfactuur_reference || "-")}
                    </td>
                    <td>{action.sheet_sync?.ok ? "ok" : action.sheet_sync?.error || "-"}</td>
                    <td>{action.email_sync?.ok ? "ok" : action.email_sync?.error || "-"}</td>
                    <td>
                      <div className="retry-actions">
                        {isEditing ? (
                          <>
                            <button
                              type="button"
                              disabled={busyActionId === `${action.id}:save`}
                              onClick={() => saveEdit(action.id)}
                            >
                              Save
                            </button>
                            <button type="button" onClick={cancelEdit}>Cancel</button>
                          </>
                        ) : canModify && (
                          <button type="button" onClick={() => startEdit(action)}>Edit</button>
                        )}
                        {!action.sheet_sync?.ok && (
                          <button
                            type="button"
                            disabled={busyActionId === `${action.id}:retry-sheet`}
                            onClick={() => retryAction(action.id, "retry-sheet")}
                          >
                            Retry sheet
                          </button>
                        )}
                        {!action.email_sync?.ok && (
                          <button
                            type="button"
                            disabled={busyActionId === `${action.id}:retry-email`}
                            onClick={() => retryAction(action.id, "retry-email")}
                          >
                            Retry email
                          </button>
                        )}
                        {canModify && !isEditing && (
                          <button
                            type="button"
                            disabled={busyActionId === `${action.id}:delete`}
                            onClick={() => deleteLocalAction(action.id)}
                          >
                            Delete local
                          </button>
                        )}
                      </div>
                    </td>
                  </tr>
                );
              })}
              {!visibleActions.length && (
                <tr>
                  <td colSpan="18">No actions were found.</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
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
  const { loading: metaLoading, data: metaData, error: metaError } = useFustMeta(Boolean(currentUser));
  const [backups, setBackups] = useState([]);
  const [backupBusy, setBackupBusy] = useState(false);
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
        <div className="form-grid">
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
          <label className="wide">
            <span>UKdocs Print spreadsheet ID</span>
            <input
              value={form.ukdocs_print_spreadsheet_id || ""}
              onChange={(event) => setForm({ ...form, ukdocs_print_spreadsheet_id: event.target.value })}
              placeholder="Spreadsheet ID for PD keuringen sendings"
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
            <span>UKdocs Print tab</span>
            <input
              value={form.ukdocs_print_sheet_name || "PD keuringen"}
              onChange={(event) => setForm({ ...form, ukdocs_print_sheet_name: event.target.value })}
              placeholder="PD keuringen"
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
          <button type="button" onClick={loadConnectionTest} disabled={connectionBusy}>
            {connectionBusy ? "Testing..." : "Test again"}
          </button>
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
                <th>Download</th>
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
