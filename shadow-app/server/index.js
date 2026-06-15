import { createReadStream, existsSync, promises as fs } from "node:fs";
import http from "node:http";
import os from "node:os";
import path from "node:path";
import crypto from "node:crypto";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const appRoot = path.resolve(__dirname, "..");
const repoRoot = path.resolve(appRoot, "..");
const cacheDir = path.resolve(process.env.SNAPPYSJAAK_CACHE_DIR || path.join(repoRoot, ".cache"));
const runDataPath = path.join(cacheDir, "run_data.json");
const syncStatusPath = path.join(cacheDir, "index_sync_status.json");
const usersPath = path.join(cacheDir, "shadow-users.json");
const fustActionsPath = path.join(cacheDir, "fust-actions.json");
const fustSettingsPath = path.join(cacheDir, "fust-settings.json");
const clockRecordsPath = path.join(cacheDir, "clock-records.json");
const fustBackupDir = path.join(cacheDir, "fust-backups");
const syncScriptPath = path.join(repoRoot, "sync_index.py");
const driveBridgePath = path.join(appRoot, "server", "drive_bridge.py");
const syncWorkerPath = path.join(appRoot, "server", "sync_worker.js");
const googleImageCacheDir = path.join(cacheDir, "shadow-google-images");
const googleRunDetailsCacheDir = path.join(cacheDir, "shadow-google-run-details");
const usersSeedPathCandidates = [
  process.env.SHADOW_USERS_SEED_PATH,
  process.platform === "win32" ? null : "/etc/secrets/shadow-users.json",
].filter(Boolean);
const staticRoot = existsSync(path.join(appRoot, "dist"))
  ? path.join(appRoot, "dist")
  : path.join(appRoot, "public");
const autoSyncOnVisit = process.env.AUTO_SYNC_ON_VISIT !== "0";
const autoSyncThrottleMs = Number(process.env.AUTO_SYNC_THROTTLE_MINUTES || 5) * 60 * 1000;
const autoSyncStartedAt = new Map();
const recentPreloadDays = Math.max(0, Number(process.env.SHADOW_PRELOAD_RECENT_DAYS || 3));
const recentPreloadMaxImagesRaw = Number(process.env.SHADOW_PRELOAD_MAX_IMAGES || 120);
const recentPreloadMaxImages = Number.isFinite(recentPreloadMaxImagesRaw)
  ? recentPreloadMaxImagesRaw
  : 120;
const googleRunDetailsCacheTtlMinutes = Math.max(0, Number(process.env.SHADOW_RUN_DETAILS_TTL_MINUTES || 720));
const recentPreloadStartedAt = new Map();
const sessions = new Map();
const sessionCookieName = "snappysjaak_shadow_session";
const allPermissions = [
  "photos:view",
  "fust:view",
  "fust:in",
  "fust:out",
  "fust:overview",
  "clock:view",
  "clock:manage",
  "users:manage",
  "settings:manage",
];
const PERMISSIONS = {
  PHOTOS_VIEW: "photos:view",
  FUST_VIEW: "fust:view",
  FUST_IN: "fust:in",
  FUST_OUT: "fust:out",
  FUST_OVERVIEW: "fust:overview",
  CLOCK_VIEW: "clock:view",
  CLOCK_MANAGE: "clock:manage",
  USERS_MANAGE: "users:manage",
  SETTINGS_MANAGE: "settings:manage",
};
const roleDefaultPermissions = {
  admin: allPermissions,
  viewer: ["photos:view"],
};
const defaultFustSettings = {
  spreadsheet_id: "",
  data_sheet_name: "Data",
  in_sheet_name: "Retour",
  out_sheet_name: "Uitgaand",
  dashboard_sheet_name: "Dashboard",
  email_recipients: [],
  smtp_host: "",
  smtp_port: 587,
  smtp_username: "",
  smtp_password: "",
  smtp_from: "",
  smtp_starttls: true,
  cmr_country_folders: {},
  cmr_fallback_folder_id: "",
  cmr_google_client_id: "",
  cmr_google_client_secret: "",
  cmr_google_refresh_token: "",
  cmr_google_connected_email: "",
  clock_spreadsheet_id: "",
  clock_employee_sheet_name: "badges",
  clock_records_sheet_name: "backup",
};

const imageExtensions = new Set([
  ".jpg",
  ".jpeg",
  ".png",
  ".gif",
  ".bmp",
  ".webp",
  ".tif",
  ".tiff",
]);

function resolvePythonCommand() {
  if (process.env.PYTHON) {
    return process.env.PYTHON;
  }
  return process.platform === "win32" ? "python" : "python3";
}

function sendJson(res, status, payload) {
  const body = JSON.stringify(payload);
  res.writeHead(status, {
    "content-type": "application/json; charset=utf-8",
    "cache-control": "no-store",
  });
  res.end(body);
}

function sendText(res, status, text) {
  res.writeHead(status, { "content-type": "text/plain; charset=utf-8" });
  res.end(text);
}

function sendUnauthorized(res) {
  sendJson(res, 401, { error: "Login required" });
}

function sendForbidden(res) {
  sendJson(res, 403, { error: "Admin access required" });
}

function parseCookies(req) {
  const cookies = {};
  const header = req.headers.cookie || "";
  for (const part of header.split(";")) {
    const [name, ...valueParts] = part.trim().split("=");
    if (!name) {
      continue;
    }
    cookies[name] = decodeURIComponent(valueParts.join("="));
  }
  return cookies;
}

function setSessionCookie(res, token) {
  res.setHeader("set-cookie", `${sessionCookieName}=${encodeURIComponent(token)}; Path=/; HttpOnly; SameSite=Lax; Max-Age=604800`);
}

function clearSessionCookie(res) {
  res.setHeader("set-cookie", `${sessionCookieName}=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0`);
}

async function readRequestJson(req, maxBytes = 1024 * 1024) {
  const chunks = [];
  let totalBytes = 0;
  for await (const chunk of req) {
    chunks.push(chunk);
    totalBytes += chunk.length;
    if (totalBytes > maxBytes) {
      throw new Error("Request body is too large");
    }
  }

  const rawBody = Buffer.concat(chunks).toString("utf8").trim();
  if (!rawBody) {
    return {};
  }
  return JSON.parse(rawBody);
}

function parseRunFolderName(folderName) {
  const parts = folderName.trim().split("_");
  const dateIndex = parts.findIndex((part) => /^\d{8}$/.test(part));
  if (dateIndex <= 0) {
    return null;
  }

  const customerRaw = parts.slice(0, dateIndex).join("_").trim();
  const runId = parts.slice(dateIndex + 1).join("_").trim() || null;
  const rawDate = parts[dateIndex];
  const year = Number(rawDate.slice(0, 4));
  const month = Number(rawDate.slice(4, 6));
  const day = Number(rawDate.slice(6, 8));
  const parsedDate = new Date(Date.UTC(year, month - 1, day));

  if (
    !customerRaw ||
    parsedDate.getUTCFullYear() !== year ||
    parsedDate.getUTCMonth() !== month - 1 ||
    parsedDate.getUTCDate() !== day
  ) {
    return null;
  }

  return {
    customer_code: customerRaw.replaceAll("_", "#"),
    run_date: `${rawDate.slice(0, 4)}-${rawDate.slice(4, 6)}-${rawDate.slice(6, 8)}`,
    run_id: runId,
  };
}

function guessMimeType(filePath) {
  const extension = path.extname(filePath).toLowerCase();
  const map = {
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
    ".gif": "image/gif",
    ".bmp": "image/bmp",
    ".webp": "image/webp",
    ".tif": "image/tiff",
    ".tiff": "image/tiff",
    ".txt": "text/plain",
    ".json": "application/json",
    ".html": "text/html",
    ".css": "text/css",
    ".js": "text/javascript",
    ".svg": "image/svg+xml",
  };
  return map[extension] || "application/octet-stream";
}

async function readJsonFile(filePath, fallback) {
  try {
    return JSON.parse(await fs.readFile(filePath, "utf8"));
  } catch {
    return fallback;
  }
}

async function writeJsonFile(filePath, payload) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(filePath, JSON.stringify(payload, null, 2), "utf8");
}

async function ensureUsersSeeded() {
  if (existsSync(usersPath)) {
    return;
  }

  for (const seedPath of usersSeedPathCandidates) {
    if (!seedPath || !existsSync(seedPath)) {
      continue;
    }

    try {
      const payload = JSON.parse(await fs.readFile(seedPath, "utf8"));
      if (!Array.isArray(payload?.users)) {
        continue;
      }
      await writeJsonFile(usersPath, payload);
      return;
    } catch {
      // Ignore invalid seed files and keep trying candidates.
    }
  }
}

function normalizePermissions(role, permissions) {
  if (role === "admin") {
    return [...roleDefaultPermissions.admin];
  }
  const defaults = roleDefaultPermissions[role] || roleDefaultPermissions.viewer;
  if (!Array.isArray(permissions)) {
    return [...defaults];
  }

  const allowed = new Set(allPermissions);
  const normalized = [...new Set(
    permissions
      .map((value) => String(value || "").trim())
      .filter((value) => allowed.has(value)),
  )];

  return normalized.length ? normalized : [...defaults];
}

function sanitizeStoredUser(user) {
  const role = user?.role === "admin" ? "admin" : "viewer";
  return {
    ...user,
    role,
    permissions: normalizePermissions(role, user?.permissions),
  };
}

function normalizeEmailRecipients(recipients) {
  if (!Array.isArray(recipients)) {
    return [];
  }

  return [...new Set(
    recipients
      .map((value) => String(value || "").trim().toLowerCase())
      .filter((value) => value.includes("@")),
  )];
}

function normalizeCmrCountryFolders(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return {};
  }
  return Object.fromEntries(
    Object.entries(value)
      .map(([country, folderId]) => [String(country || "").trim().toUpperCase(), String(folderId || "").trim()])
      .filter(([country, folderId]) => country && folderId),
  );
}

function normalizeCmrInfo(value) {
  const status = ["uploaded", "skipped", "failed"].includes(value?.status) ? value.status : "missing";
  return {
    status,
    file_id: String(value?.file_id || ""),
    file_name: String(value?.file_name || ""),
    web_link: String(value?.web_link || ""),
    mime_type: String(value?.mime_type || ""),
    folder_id: String(value?.folder_id || ""),
    error: String(value?.error || ""),
    uploaded_at: String(value?.uploaded_at || ""),
    uploaded_by: String(value?.uploaded_by || ""),
  };
}

function normalizeFustSettings(settings) {
  const smtpPort = Number(settings?.smtp_port);
  return {
    spreadsheet_id: String(settings?.spreadsheet_id || "").trim(),
    data_sheet_name: String(settings?.data_sheet_name || defaultFustSettings.data_sheet_name).trim() || defaultFustSettings.data_sheet_name,
    in_sheet_name: String(settings?.in_sheet_name || defaultFustSettings.in_sheet_name).trim() || defaultFustSettings.in_sheet_name,
    out_sheet_name: String(settings?.out_sheet_name || defaultFustSettings.out_sheet_name).trim() || defaultFustSettings.out_sheet_name,
    dashboard_sheet_name: String(settings?.dashboard_sheet_name || defaultFustSettings.dashboard_sheet_name).trim() || defaultFustSettings.dashboard_sheet_name,
    email_recipients: normalizeEmailRecipients(settings?.email_recipients),
    smtp_host: String(settings?.smtp_host || "").trim(),
    smtp_port: Number.isFinite(smtpPort) && smtpPort > 0 ? smtpPort : defaultFustSettings.smtp_port,
    smtp_username: String(settings?.smtp_username || "").trim(),
    smtp_password: String(settings?.smtp_password || ""),
    smtp_from: String(settings?.smtp_from || "").trim(),
    smtp_starttls: settings?.smtp_starttls === false || settings?.smtp_starttls === "0" || settings?.smtp_starttls === "false"
      ? false
      : true,
    cmr_country_folders: normalizeCmrCountryFolders(settings?.cmr_country_folders),
    cmr_fallback_folder_id: String(settings?.cmr_fallback_folder_id || "").trim(),
    cmr_google_client_id: String(settings?.cmr_google_client_id || "").trim(),
    cmr_google_client_secret: String(settings?.cmr_google_client_secret || ""),
    cmr_google_refresh_token: String(settings?.cmr_google_refresh_token || ""),
    cmr_google_connected_email: String(settings?.cmr_google_connected_email || "").trim(),
    clock_spreadsheet_id: String(settings?.clock_spreadsheet_id || "").trim(),
    clock_employee_sheet_name: String(settings?.clock_employee_sheet_name || defaultFustSettings.clock_employee_sheet_name).trim() || defaultFustSettings.clock_employee_sheet_name,
    clock_records_sheet_name: String(settings?.clock_records_sheet_name || defaultFustSettings.clock_records_sheet_name).trim() || defaultFustSettings.clock_records_sheet_name,
  };
}

async function readFustSettings() {
  const payload = await readJsonFile(fustSettingsPath, defaultFustSettings);
  return normalizeFustSettings(payload);
}

async function writeFustSettings(settings) {
  await writeJsonFile(fustSettingsPath, normalizeFustSettings(settings));
}

function normalizeFustAction(action) {
  return {
    id: String(action?.id || ""),
    type: action?.type === "OUT" ? "OUT" : "IN",
    action_date: String(action?.action_date || ""),
    week: Number.isFinite(Number(action?.week)) ? Number(action.week) : null,
    day_name: String(action?.day_name || ""),
    country: String(action?.country || "").trim(),
    customer_name: String(action?.customer_name || "").trim(),
    customer_code: String(action?.customer_code || "").trim(),
    connect_name: String(action?.connect_name || "").trim(),
    remark: String(action?.remark || "").trim(),
    metrics: {
      dc: Number(action?.metrics?.dc || 0),
      cctag: Number(action?.metrics?.cctag || 0),
      dcs: Number(action?.metrics?.dcs || 0),
      dco: Number(action?.metrics?.dco || 0),
      pal: Number(action?.metrics?.pal || 0),
      vk: Number(action?.metrics?.vk || 0),
    },
    created_by: String(action?.created_by || ""),
    created_at: String(action?.created_at || ""),
    sheet_sync: action?.sheet_sync || { ok: false, target_sheets: [], error: "Not attempted" },
    email_sync: action?.email_sync || { ok: false, recipients: [], error: "Not attempted" },
    cmr: normalizeCmrInfo(action?.cmr),
    fustbon: normalizeCmrInfo(action?.fustbon),
  };
}

async function readFustActions() {
  const payload = await readJsonFile(fustActionsPath, { actions: [] });
  return Array.isArray(payload?.actions) ? payload.actions.map(normalizeFustAction) : [];
}

async function writeFustActions(actions) {
  await writeJsonFile(fustActionsPath, { actions: actions.map(normalizeFustAction) });
}

function normalizeClockEmployee(employee) {
  return {
    tbnr: String(employee?.tbnr || employee?.TBNR || "").trim().toUpperCase(),
    type: String(employee?.type || employee?.employee_type || "").trim(),
    name: String(employee?.name || employee?.Name || "").trim(),
    active: employee?.active === false ? false : true,
  };
}

function buildClockEmployeesFromSheetRows(rows) {
  if (!Array.isArray(rows) || rows.length < 2) {
    return { employees: [], headers: [], raw_row_count: Array.isArray(rows) ? rows.length : 0 };
  }

  const headers = rows[0].map(normalizeHeader);
  const tbnrIndex = firstMatchingIndex(headers, ["tbnr", "badge", "badge number", "badge nummer", "code", "qr", "qr code"]);
  const typeIndex = firstMatchingIndex(headers, ["type", "employee type", "contract", "soort"]);
  const nameIndex = firstMatchingIndex(headers, ["name", "naam", "employee", "werknemer", "medewerker"]);
  const activeIndex = firstMatchingIndex(headers, ["active", "actief", "enabled", "status"]);
  const sourceRows = tbnrIndex >= 0 || nameIndex >= 0 ? rows.slice(1) : rows;

  const employees = sourceRows
    .map((row) => {
      const activeValue = rowValue(row, activeIndex);
      return normalizeClockEmployee({
        tbnr: rowValue(row, tbnrIndex >= 0 ? tbnrIndex : 0),
        type: rowValue(row, typeIndex >= 0 ? typeIndex : 1),
        name: rowValue(row, nameIndex >= 0 ? nameIndex : 2),
        active: activeIndex < 0 ? true : !["0", "false", "nee", "no", "inactive"].includes(activeValue.toLowerCase()),
      });
    })
    .filter((employee) => employee.tbnr && employee.name && employee.active)
    .sort((left, right) => left.name.localeCompare(right.name) || left.tbnr.localeCompare(right.tbnr));

  return { employees, headers, raw_row_count: rows.length };
}

function normalizeClockRecord(record) {
  const direction = String(record?.direction || "").trim().toUpperCase() === "OUT" ? "OUT" : "IN";
  const actionDate = String(record?.action_date || record?.date || localDateIso()).slice(0, 10);
  const actionTime = String(record?.action_time || record?.time || "").trim();
  return {
    id: String(record?.id || crypto.randomUUID()),
    action_date: actionDate,
    action_time: actionTime || new Date().toLocaleTimeString("nl-NL", { hour12: false }),
    timestamp: String(record?.timestamp || `${actionDate}T${actionTime || "00:00:00"}`),
    tbnr: String(record?.tbnr || "").trim().toUpperCase(),
    name: String(record?.name || "").trim(),
    employee_type: String(record?.employee_type || record?.type || "").trim(),
    direction,
    source: String(record?.source || "scanner").trim() || "scanner",
    created_by: String(record?.created_by || ""),
    created_at: String(record?.created_at || new Date().toISOString()),
    updated_by: String(record?.updated_by || ""),
    updated_at: String(record?.updated_at || ""),
    sheet_sync: record?.sheet_sync || { ok: false, target_sheets: [], error: "Not attempted" },
  };
}

async function readClockRecords() {
  const payload = await readJsonFile(clockRecordsPath, { records: [] });
  return Array.isArray(payload?.records) ? payload.records.map(normalizeClockRecord) : [];
}

async function writeClockRecords(records) {
  await writeJsonFile(clockRecordsPath, { records: records.map(normalizeClockRecord) });
}

function clockTimestampMs(record) {
  const parsed = new Date(`${record.action_date}T${record.action_time || "00:00:00"}`);
  return Number.isNaN(parsed.getTime()) ? null : parsed.getTime();
}

function formatWorkedMinutes(minutes) {
  if (!Number.isFinite(minutes) || minutes <= 0) {
    return "";
  }
  const rounded = Math.round(minutes);
  const hours = Math.floor(rounded / 60);
  const remainder = rounded % 60;
  return `${hours}:${String(remainder).padStart(2, "0")}`;
}

function clockWorkedMinutes(record, allRecords = []) {
  if (record.direction !== "OUT") {
    return 0;
  }
  const outMs = clockTimestampMs(record);
  if (!outMs) {
    return 0;
  }
  const previousIn = allRecords
    .filter((item) => item.action_date === record.action_date && item.tbnr === record.tbnr && item.direction === "IN")
    .map((item) => ({ item, ms: clockTimestampMs(item) }))
    .filter(({ ms }) => ms !== null && ms <= outMs)
    .sort((left, right) => right.ms - left.ms)[0];
  if (!previousIn) {
    return 0;
  }
  return Math.max(0, (outMs - previousIn.ms) / 60000);
}

function addClockWorkedDurations(records) {
  const normalized = records.map(normalizeClockRecord);
  return normalized.map((record) => {
    const workedMinutes = clockWorkedMinutes(record, normalized);
    return {
      ...record,
      worked_minutes: Math.round(workedMinutes),
      worked_time: formatWorkedMinutes(workedMinutes),
    };
  });
}

function previousClockInForOut(record, allRecords = []) {
  if (record.direction !== "OUT") {
    return null;
  }
  const outMs = clockTimestampMs(record);
  if (!outMs) {
    return null;
  }
  return allRecords
    .filter((item) => item.action_date === record.action_date && item.tbnr === record.tbnr && item.direction === "IN")
    .map((item) => ({ item, ms: clockTimestampMs(item) }))
    .filter(({ ms }) => ms !== null && ms <= outMs)
    .sort((left, right) => right.ms - left.ms)[0]?.item || null;
}

function clockSessionRow(inRecord, outRecord = null, allRecords = []) {
  const base = inRecord?.direction === "IN" ? inRecord : previousClockInForOut(outRecord, allRecords);
  const first = base || inRecord || outRecord;
  const workedMinutes = outRecord ? clockWorkedMinutes(outRecord, allRecords) : 0;
  return [
    first?.action_date || outRecord?.action_date || "",
    first?.tbnr || outRecord?.tbnr || "",
    first?.name || outRecord?.name || "",
    first?.employee_type || outRecord?.employee_type || "",
    base?.action_time || "",
    outRecord?.action_time || "",
    formatWorkedMinutes(workedMinutes),
    outRecord?.source || first?.source || "",
    first?.id || outRecord?.id || "",
    first?.created_by || outRecord?.created_by || "",
    first?.created_at || outRecord?.created_at || "",
  ];
}

function clockSessionRows(records) {
  const sorted = addClockWorkedDurations(records)
    .sort((left, right) => `${left.action_date}T${left.action_time}`.localeCompare(`${right.action_date}T${right.action_time}`));
  const openByEmployeeDay = new Map();
  const rows = [];
  for (const record of sorted) {
    const key = `${record.action_date}__${record.tbnr}`;
    if (record.direction === "IN") {
      openByEmployeeDay.set(key, record);
      rows.push({ inRecord: record, outRecord: null, row: clockSessionRow(record, null, sorted) });
      continue;
    }
    const openIn = openByEmployeeDay.get(key);
    if (openIn) {
      const existing = rows.find((item) => item.inRecord?.id === openIn.id);
      if (existing) {
        existing.outRecord = record;
        existing.row = clockSessionRow(openIn, record, sorted);
      } else {
        rows.push({ inRecord: openIn, outRecord: record, row: clockSessionRow(openIn, record, sorted) });
      }
      openByEmployeeDay.delete(key);
    } else {
      rows.push({ inRecord: null, outRecord: record, row: clockSessionRow(null, record, sorted) });
    }
  }
  return rows;
}

async function syncClockRecordToSheets(record, settings, allRecords = []) {
  if (!settings.clock_spreadsheet_id) {
    return { ok: false, target_sheets: [], error: "Clock spreadsheet ID is not configured" };
  }
  if (!settings.clock_records_sheet_name) {
    return { ok: false, target_sheets: [], error: "Clock records sheet is not configured" };
  }

  const previousIn = previousClockInForOut(record, allRecords);
  const previousRowNumber = Number(previousIn?.sheet_sync?.row_number || 0);
  const sessionRow = record.direction === "OUT"
    ? clockSessionRow(previousIn, record, allRecords)
    : clockSessionRow(record, null, allRecords);
  const output = record.direction === "OUT" && previousRowNumber >= 2
    ? await writeSheetRowAt(settings.clock_spreadsheet_id, settings.clock_records_sheet_name, previousRowNumber, sessionRow)
    : await writeSheetRowToFirstEmpty(settings.clock_spreadsheet_id, settings.clock_records_sheet_name, sessionRow);

  return {
    ok: true,
    target_sheets: [settings.clock_records_sheet_name],
    error: "",
    synced_at: new Date().toISOString(),
    row_number: Number(output?.row_number || previousRowNumber || 0),
  };
}

function nextClockDirection(records, tbnr, actionDate) {
  const count = records.filter((record) => record.action_date === actionDate && record.tbnr.toUpperCase() === tbnr.toUpperCase()).length;
  return count % 2 === 0 ? "IN" : "OUT";
}

function createClockRecord(employee, direction, actionDate, actionTime, source, username) {
  return normalizeClockRecord({
    id: crypto.randomUUID(),
    action_date: actionDate,
    action_time: actionTime,
    timestamp: `${actionDate}T${actionTime}`,
    tbnr: employee.tbnr,
    name: employee.name,
    employee_type: employee.type,
    direction,
    source,
    created_by: username,
    created_at: new Date().toISOString(),
    sheet_sync: { ok: false, target_sheets: [], error: "Pending" },
  });
}

function filterClockRecords(records, date) {
  const activeDate = String(date || localDateIso()).slice(0, 10);
  return addClockWorkedDurations(records)
    .filter((record) => record.action_date === activeDate)
    .sort((left, right) => `${right.action_date}T${right.action_time}`.localeCompare(`${left.action_date}T${left.action_time}`));
}

function clockExportText(records) {
  const rows = [
    ["Date", "TBNR", "Name", "Type", "IN", "OUT", "Worked time", "Source", "ID", "Created by", "Created at"],
    ...clockSessionRows(records).map((session) => session.row),
  ];
  return rows.map((row) => row.map((value) => String(value || "").replaceAll("\t", " ")).join("\t")).join("\n");
}

function normalizeClockTime(value) {
  const raw = String(value || "").trim();
  const match = raw.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!match) {
    return "";
  }
  return `${String(match[1]).padStart(2, "0")}:${match[2]}:${match[3] || "00"}`;
}

function parseClockBackupRows(rows, targetDate, username) {
  if (!Array.isArray(rows) || rows.length < 2) {
    return [];
  }

  const headers = rows[0].map(normalizeHeader);
  const dateIndex = firstMatchingIndex(headers, ["date", "datum"]);
  const tbnrIndex = firstMatchingIndex(headers, ["tbnr", "badge", "badge number", "badge nummer"]);
  const nameIndex = firstMatchingIndex(headers, ["name", "naam"]);
  const typeIndex = firstMatchingIndex(headers, ["type", "employee type"]);
  const inIndex = firstMatchingIndex(headers, ["in", "start"]);
  const outIndex = firstMatchingIndex(headers, ["out", "uit", "finish", "einde"]);
  const sourceIndex = firstMatchingIndex(headers, ["source", "bron"]);
  const idIndex = firstMatchingIndex(headers, ["id"]);
  const createdByIndex = firstMatchingIndex(headers, ["created by", "gemaakt door"]);
  const createdAtIndex = firstMatchingIndex(headers, ["created at", "gemaakt op"]);

  const imported = [];
  for (const [offset, row] of rows.slice(1).entries()) {
    const rowNumber = offset + 2;
    const actionDate = parseSheetDateToIso(rowValue(row, dateIndex >= 0 ? dateIndex : 0)).slice(0, 10);
    if (actionDate !== targetDate) {
      continue;
    }

    const tbnr = rowValue(row, tbnrIndex >= 0 ? tbnrIndex : 1).toUpperCase();
    const name = rowValue(row, nameIndex >= 0 ? nameIndex : 2);
    const employeeType = rowValue(row, typeIndex >= 0 ? typeIndex : 3);
    const inTime = normalizeClockTime(rowValue(row, inIndex >= 0 ? inIndex : 4));
    const outTime = normalizeClockTime(rowValue(row, outIndex >= 0 ? outIndex : 5));
    const source = rowValue(row, sourceIndex >= 0 ? sourceIndex : 7) || "backup";
    const baseId = rowValue(row, idIndex >= 0 ? idIndex : 8) || crypto.randomUUID();
    const createdBy = rowValue(row, createdByIndex >= 0 ? createdByIndex : 9) || username;
    const createdAt = rowValue(row, createdAtIndex >= 0 ? createdAtIndex : 10) || new Date().toISOString();

    if (!tbnr || !name) {
      continue;
    }

    if (inTime) {
      imported.push(normalizeClockRecord({
        id: baseId,
        action_date: actionDate,
        action_time: inTime,
        timestamp: `${actionDate}T${inTime}`,
        tbnr,
        name,
        employee_type: employeeType,
        direction: "IN",
        source,
        created_by: createdBy,
        created_at: createdAt,
        sheet_sync: { ok: true, target_sheets: [], error: "", row_number: rowNumber },
      }));
    }

    if (outTime) {
      imported.push(normalizeClockRecord({
        id: `${baseId}-out`,
        action_date: actionDate,
        action_time: outTime,
        timestamp: `${actionDate}T${outTime}`,
        tbnr,
        name,
        employee_type: employeeType,
        direction: "OUT",
        source,
        created_by: createdBy,
        created_at: createdAt,
        sheet_sync: { ok: true, target_sheets: [], error: "", row_number: rowNumber },
      }));
    }
  }

  return imported.sort((left, right) => `${left.action_date}T${left.action_time}`.localeCompare(`${right.action_date}T${right.action_time}`));
}

async function createFustBackupSnapshot(createdBy = "system") {
  const [settings, actions] = await Promise.all([readFustSettings(), readFustActions()]);
  const timestamp = new Date().toISOString().replaceAll(":", "-");
  const filename = `fust-backup-${timestamp}.json`;
  const filePath = path.join(fustBackupDir, filename);
  const payload = {
    created_at: new Date().toISOString(),
    created_by: createdBy,
    settings,
    actions,
  };
  await writeJsonFile(filePath, payload);
  return {
    filename,
    created_at: payload.created_at,
    created_by: createdBy,
    action_count: actions.length,
    size_bytes: Buffer.byteLength(JSON.stringify(payload, null, 2), "utf8"),
  };
}

async function listFustBackups() {
  await fs.mkdir(fustBackupDir, { recursive: true });
  const entries = await fs.readdir(fustBackupDir, { withFileTypes: true });
  const backups = [];
  for (const entry of entries) {
    if (!entry.isFile() || path.extname(entry.name).toLowerCase() !== ".json") {
      continue;
    }
    const filePath = path.join(fustBackupDir, entry.name);
    const stat = await fs.stat(filePath).catch(() => null);
    backups.push({
      filename: entry.name,
      created_at: stat?.mtime?.toISOString() || "",
      size_bytes: stat?.size || 0,
      download_path: `/api/fust/backups/${encodeURIComponent(entry.name)}`,
    });
  }
  backups.sort((left, right) => String(right.created_at).localeCompare(String(left.created_at)));
  return backups;
}

function isoDateForDisplay(value) {
  if (!value) {
    return "";
  }
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) {
    return String(value);
  }
  const day = String(parsed.getDate()).padStart(2, "0");
  const month = String(parsed.getMonth() + 1).padStart(2, "0");
  const year = String(parsed.getFullYear()).slice(-2);
  return `${day}-${month}-${year}`;
}

function weekNumberForDate(dateString) {
  const parsed = new Date(`${dateString}T00:00:00`);
  if (Number.isNaN(parsed.getTime())) {
    return null;
  }
  const utc = new Date(Date.UTC(parsed.getFullYear(), parsed.getMonth(), parsed.getDate()));
  const dayNum = utc.getUTCDay() || 7;
  utc.setUTCDate(utc.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(utc.getUTCFullYear(), 0, 1));
  return Math.ceil((((utc - yearStart) / 86400000) + 1) / 7);
}

function weekdayNameForDate(dateString) {
  const parsed = new Date(`${dateString}T00:00:00`);
  if (Number.isNaN(parsed.getTime())) {
    return "";
  }
  return parsed.toLocaleDateString("nl-NL", { weekday: "long" }).toLowerCase();
}

function requirePermission(res, requestUser, permission) {
  const permissions = normalizePermissions(requestUser?.role, requestUser?.permissions);
  if (!permissions.includes(permission)) {
    sendJson(res, 403, { error: "You do not have access to this action" });
    return false;
  }
  return true;
}

function normalizeHeader(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function firstMatchingIndex(headers, aliases) {
  return headers.findIndex((header) => aliases.includes(header));
}

function rowValue(row, index) {
  if (index < 0 || !Array.isArray(row)) {
    return "";
  }
  return String(row[index] || "").trim();
}

function buildFustMetaFromSheetRows(rows) {
  if (!Array.isArray(rows) || rows.length < 2) {
    return { countries: [], records: [], headers: [], raw_row_count: Array.isArray(rows) ? rows.length : 0 };
  }

  const headers = rows[0].map(normalizeHeader);
  // Primary expected Data-tab headers:
  // klantnaam | Country | klantcode connect
  const countryIndex = firstMatchingIndex(headers, ["country", "land", "co", "country code", "land code"]);
  const customerIndex = firstMatchingIndex(headers, ["klantnaam", "customer", "carrier", "cust transport", "transport", "customer name"]);
  const connectIndex = firstMatchingIndex(headers, ["klantcode connect", "connect", "connect name", "connect code", "klantcode connector"]);
  const customerCodeIndex = firstMatchingIndex(headers, ["customer code", "klantcode", "code", "customer id"]);
  const activeIndex = firstMatchingIndex(headers, ["active", "actief", "enabled", "status"]);
  const hasHeaderRow = countryIndex >= 0 || customerIndex >= 0 || connectIndex >= 0;

  const sourceRows = hasHeaderRow ? rows.slice(1) : rows;
  const fallbackCountryIndex = 5;
  const fallbackCustomerIndex = 4;
  const fallbackConnectIndex = 6;

  const records = sourceRows
    .map((row) => {
      const country = rowValue(row, countryIndex >= 0 ? countryIndex : fallbackCountryIndex);
      const customerName = rowValue(row, customerIndex >= 0 ? customerIndex : fallbackCustomerIndex);
      const connectName = rowValue(row, connectIndex >= 0 ? connectIndex : fallbackConnectIndex);
      const customerCode = rowValue(row, customerCodeIndex >= 0 ? customerCodeIndex : (connectIndex >= 0 ? connectIndex : fallbackConnectIndex));
      const activeValue = rowValue(row, activeIndex);

      return {
        country,
        customer_name: customerName,
        connect_name: connectName || customerCode,
        customer_code: customerCode || connectName,
        active: activeIndex < 0 ? true : !["0", "false", "nee", "no", "inactive"].includes(activeValue.toLowerCase()),
      };
    })
    .filter((row) => row.country && row.customer_name && row.active);

  const countries = [...new Set(records.map((row) => row.country))].sort((left, right) => left.localeCompare(right));
  return {
    countries,
    records,
    headers,
    raw_row_count: rows.length,
  };
}

function buildOverview(actions) {
  const grouped = new Map();
  for (const action of actions) {
    const key = `${action.week ?? ""}__${action.country}__${action.customer_name}`;
    if (!grouped.has(key)) {
      grouped.set(key, {
        week: action.week ?? null,
        country: action.country,
        customer_name: action.customer_name,
        connect_names: new Set(),
        in: { dc: 0, cctag: 0, dcs: 0, dco: 0, pal: 0, vk: 0 },
        out: { dc: 0, cctag: 0, dcs: 0, dco: 0, pal: 0, vk: 0 },
      });
    }
    const entry = grouped.get(key);
    entry.connect_names.add(action.connect_name);
    const target = action.type === "OUT" ? entry.out : entry.in;
    for (const metric of Object.keys(target)) {
      target[metric] += Number(action.metrics?.[metric] || 0);
    }
  }

  return [...grouped.values()].map((entry) => ({
    week: entry.week,
    country: entry.country,
    customer_name: entry.customer_name,
    connect_names: [...entry.connect_names].filter(Boolean).sort((left, right) => left.localeCompare(right)),
    in: entry.in,
    out: entry.out,
    balance: {
      dc: entry.in.dc - entry.out.dc,
      cctag: entry.in.cctag - entry.out.cctag,
      dcs: entry.in.dcs - entry.out.dcs,
      dco: entry.in.dco - entry.out.dco,
      pal: entry.in.pal - entry.out.pal,
      vk: entry.in.vk - entry.out.vk,
    },
  }));
}

function normalizeNumber(value) {
  const raw = String(value || "").trim().replace(",", ".");
  const parsed = Number(raw);
  return Number.isFinite(parsed) ? parsed : 0;
}

function parseSheetDateToIso(value) {
  const raw = String(value || "").trim();
  const match = raw.match(/^(\d{1,2})-(\d{1,2})-(\d{2}|\d{4})$/);
  if (!match) {
    return raw;
  }
  const [, day, month, year] = match;
  const fullYear = year.length === 2 ? `20${year}` : year;
  return `${fullYear}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function buildActionSignature(action) {
  return [
    action.type,
    action.action_date,
    action.customer_name,
    action.country,
    action.customer_code || action.connect_name,
    action.remark,
    action.metrics?.dc || 0,
    action.metrics?.cctag || 0,
    action.metrics?.dcs || 0,
    action.metrics?.dco || 0,
    action.metrics?.pal || 0,
    action.metrics?.vk || 0,
  ].join("|");
}

function parseDashboardSheetRows(rows) {
  if (!Array.isArray(rows) || rows.length < 2) {
    return [];
  }

  const headers = rows[0].map(normalizeHeader);
  const hasHeaderRow = headers.includes("richting") || headers.includes("klantnaam");
  const sourceRows = hasHeaderRow ? rows.slice(1) : rows;

  return sourceRows
    .map((row, index) => normalizeFustAction({
      id: `sheet-${index + 1}`,
      type: rowValue(row, 0).toLowerCase() === "uitgaand" ? "OUT" : "IN",
      day_name: rowValue(row, 1),
      action_date: parseSheetDateToIso(rowValue(row, 2)),
      week: rowValue(row, 3) ? Number(rowValue(row, 3)) : null,
      customer_name: rowValue(row, 4),
      country: rowValue(row, 5),
      customer_code: rowValue(row, 6),
      connect_name: rowValue(row, 6),
      remark: rowValue(row, 7),
      metrics: {
        dc: normalizeNumber(rowValue(row, 8)),
        cctag: normalizeNumber(rowValue(row, 9)),
        dcs: normalizeNumber(rowValue(row, 10)),
        dco: normalizeNumber(rowValue(row, 11)),
        pal: normalizeNumber(rowValue(row, 12)),
        vk: normalizeNumber(rowValue(row, 13)),
      },
      created_by: "spreadsheet",
      created_at: "",
      sheet_sync: { ok: true, target_sheets: ["Dashboard"], error: "" },
      email_sync: { ok: true, recipients: [], error: "" },
    }))
    .filter((action) => action.customer_name && action.country);
}

function parseRegistrySheetRows(rows, type) {
  if (!Array.isArray(rows) || rows.length < 2) {
    return [];
  }

  const headers = rows[0].map(normalizeHeader);
  const hasHeaderRow = headers.includes("klantnaam") || headers.includes("dag");
  const sourceRows = hasHeaderRow ? rows.slice(1) : rows;

  return sourceRows
    .map((row, index) => normalizeFustAction({
      id: `${type.toLowerCase()}-sheet-${index + 1}`,
      type,
      day_name: rowValue(row, 0),
      action_date: parseSheetDateToIso(rowValue(row, 1)),
      week: rowValue(row, 2) ? Number(rowValue(row, 2)) : null,
      customer_name: rowValue(row, 3),
      country: rowValue(row, 4),
      customer_code: rowValue(row, 5),
      connect_name: rowValue(row, 5),
      remark: rowValue(row, 6),
      metrics: {
        dc: normalizeNumber(rowValue(row, 7)),
        cctag: normalizeNumber(rowValue(row, 8)),
        dcs: normalizeNumber(rowValue(row, 9)),
        dco: normalizeNumber(rowValue(row, 10)),
        pal: normalizeNumber(rowValue(row, 11)),
        vk: normalizeNumber(rowValue(row, 12)),
      },
      created_by: "spreadsheet",
      created_at: "",
      sheet_sync: { ok: true, target_sheets: [type === "OUT" ? "Uitgaand" : "Retour"], error: "" },
      email_sync: { ok: true, recipients: [], error: "" },
    }))
    .filter((action) => action.customer_name && action.country);
}

function fustSheetRow(action) {
  return [
    action.day_name,
    isoDateForDisplay(action.action_date),
    action.week ?? "",
    action.customer_name,
    action.country,
    action.customer_code || action.connect_name,
    action.remark,
    action.metrics.dc || "",
    action.metrics.cctag || "",
    action.metrics.dcs || "",
    action.metrics.dco || "",
    action.metrics.pal || "",
    action.metrics.vk || "",
    "",
    "",
    "",
    "",
  ];
}

function fustDashboardRow(action) {
  return [
    action.type === "OUT" ? "uitgaand" : "retour",
    action.day_name,
    isoDateForDisplay(action.action_date),
    action.week ?? "",
    action.customer_name,
    action.country,
    action.customer_code || action.connect_name,
    action.remark,
    action.metrics.dc || "",
    action.metrics.cctag || "",
    action.metrics.dcs || "",
    action.metrics.dco || "",
    action.metrics.pal || "",
    action.metrics.vk || "",
    "",
    "",
    "",
    "",
  ];
}

async function readUsers() {
  await ensureUsersSeeded();
  const payload = await readJsonFile(usersPath, { users: [] });
  return Array.isArray(payload?.users) ? payload.users.map(sanitizeStoredUser) : [];
}

async function writeUsers(users) {
  await writeJsonFile(usersPath, { users });
}

function publicUser(user) {
  const normalizedUser = sanitizeStoredUser(user);
  return {
    username: normalizedUser.username,
    role: normalizedUser.role,
    permissions: normalizedUser.permissions,
    created_at: normalizedUser.created_at,
  };
}

function hashPassword(password, salt = crypto.randomBytes(16).toString("hex")) {
  const hash = crypto.scryptSync(String(password), salt, 64).toString("hex");
  return `${salt}:${hash}`;
}

function verifyPassword(password, storedHash) {
  const [salt, expectedHash] = String(storedHash || "").split(":");
  if (!salt || !expectedHash) {
    return false;
  }
  const actualHash = crypto.scryptSync(String(password), salt, 64);
  const expectedBuffer = Buffer.from(expectedHash, "hex");
  return expectedBuffer.length === actualHash.length && crypto.timingSafeEqual(expectedBuffer, actualHash);
}

async function getRequestUser(req) {
  const token = parseCookies(req)[sessionCookieName];
  if (!token) {
    return null;
  }

  const session = sessions.get(token);
  if (!session || session.expires_at < Date.now()) {
    sessions.delete(token);
    return null;
  }

  const users = await readUsers();
  const user = users.find((item) => item.username === session.username);
  if (!user) {
    sessions.delete(token);
    return null;
  }

  session.expires_at = Date.now() + 7 * 24 * 60 * 60 * 1000;
  return publicUser(user);
}

function createSession(res, user) {
  const token = crypto.randomBytes(32).toString("hex");
  sessions.set(token, {
    username: user.username,
    expires_at: Date.now() + 7 * 24 * 60 * 60 * 1000,
  });
  setSessionCookie(res, token);
}

function destroySession(req, res) {
  const token = parseCookies(req)[sessionCookieName];
  if (token) {
    sessions.delete(token);
  }
  clearSessionCookie(res);
}

async function readRunData() {
  const payload = await readJsonFile(runDataPath, null);
  if (!payload || !Array.isArray(payload.runs)) {
    return { runs: [], parse_errors: [], generated_at: null, cache_missing: true };
  }

  return {
    runs: payload.runs,
    parse_errors: Array.isArray(payload.parse_errors) ? payload.parse_errors : [],
    generated_at: payload.generated_at || null,
    cache_missing: false,
  };
}

function localDateIso() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

async function isSyncRunning() {
  const status = await readJsonFile(syncStatusPath, {});
  return status?.state === "running";
}

function summarizeBridgeError(rawError) {
  const text = String(rawError || "").trim();
  if (!text) {
    return text;
  }

  const lines = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  for (let index = lines.length - 1; index >= 0; index -= 1) {
    const line = lines[index];
    if (/^(RuntimeError|ValueError|Exception|Error|HttpError)\b/.test(line)) {
      return line;
    }
  }

  return lines.at(-1) || text;
}

function runPythonBridge(args, input = "") {
  return new Promise((resolve, reject) => {
    const child = spawn(resolvePythonCommand(), [driveBridgePath, ...args], {
      cwd: repoRoot,
      windowsHide: true,
    });

    const stdout = [];
    const stderr = [];
    child.stdout.on("data", (chunk) => stdout.push(chunk));
    child.stderr.on("data", (chunk) => stderr.push(chunk));
    child.on("error", reject);
    child.on("close", (code) => {
      const output = Buffer.concat(stdout);
      if (code === 0) {
        resolve(output);
        return;
      }

      reject(new Error(summarizeBridgeError(Buffer.concat(stderr).toString("utf8")) || `Python bridge exited with ${code}`));
    });

    if (input) {
      child.stdin.write(input);
    }
    child.stdin.end();
  });
}

async function loadSheetRows(spreadsheetId, sheetName) {
  if (!spreadsheetId || !sheetName) {
    return [];
  }

  const output = await runPythonBridge([
    "sheets-read",
    "--spreadsheet-id",
    spreadsheetId,
    "--sheet-name",
    sheetName,
  ]);
  const payload = JSON.parse(output.toString("utf8"));
  return Array.isArray(payload?.values) ? payload.values : [];
}

async function loadFustSheetRows(settings) {
  return loadSheetRows(settings.spreadsheet_id, settings.data_sheet_name);
}

async function loadServiceAccountInfo() {
  const output = await runPythonBridge(["service-account-info"]);
  const payload = JSON.parse(output.toString("utf8"));
  return {
    client_email: String(payload?.client_email || ""),
    project_id: String(payload?.project_id || ""),
  };
}

async function writeSheetRowToFirstEmpty(spreadsheetId, sheetName, row) {
  const output = await runPythonBridge(
    ["sheets-write-first-empty", "--spreadsheet-id", spreadsheetId, "--sheet-name", sheetName],
    JSON.stringify({ row }),
  );
  return JSON.parse(output.toString("utf8"));
}

async function writeSheetRowAt(spreadsheetId, sheetName, rowNumber, row) {
  const output = await runPythonBridge(
    ["sheets-write-row", "--spreadsheet-id", spreadsheetId, "--sheet-name", sheetName, "--row-number", String(rowNumber)],
    JSON.stringify({ row }),
  );
  return JSON.parse(output.toString("utf8"));
}

function sanitizeDriveName(value) {
  return String(value || "unknown")
    .trim()
    .replace(/[\\/:*?"<>|#%{}]/g, "-")
    .replace(/\s+/g, " ")
    .slice(0, 120) || "unknown";
}

function safeExtension(filename, mimeType) {
  const extension = path.extname(String(filename || "")).toLowerCase();
  if (extension && extension.length <= 10) {
    return extension;
  }
  if (mimeType === "application/pdf") {
    return ".pdf";
  }
  if (String(mimeType || "").includes("png")) {
    return ".png";
  }
  if (String(mimeType || "").includes("webp")) {
    return ".webp";
  }
  return ".jpg";
}

function cmrTargetFolderId(settings, action) {
  const country = String(action.country || "").trim().toUpperCase();
  return settings.cmr_country_folders?.[country] || settings.cmr_fallback_folder_id || "";
}

function buildCmrFilename(action, originalName, mimeType) {
  const extension = safeExtension(originalName, mimeType);
  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  return `${sanitizeDriveName(action.type)}-${sanitizeDriveName(action.action_date)}-${sanitizeDriveName(action.country)}-${sanitizeDriveName(action.customer_name)}-${stamp}${extension}`;
}

async function uploadFustDocumentToDrive(action, settings, filePayload, documentKind) {
  const countryFolderId = cmrTargetFolderId(settings, action);
  const documentLabel = documentKind === "fustbon" ? "Fustbon" : "CMR";
  if (!countryFolderId) {
    throw new Error(`No ${documentLabel} folder configured for ${action.country}`);
  }

  const filename = buildCmrFilename(action, filePayload.name, filePayload.type);
  const output = await runPythonBridge(
    ["drive-upload-cmr"],
    JSON.stringify({
      country_folder_id: countryFolderId,
      folder_path: [
        `${sanitizeDriveName(action.customer_name)} ${documentLabel}`,
        String(new Date(action.action_date || localDateIso()).getFullYear()),
        `Week ${action.week ?? "unknown"}`,
      ],
      filename,
      mime_type: filePayload.type || "application/octet-stream",
      content_base64: filePayload.content_base64,
      oauth: settings.cmr_google_refresh_token ? {
        client_id: settings.cmr_google_client_id,
        client_secret: settings.cmr_google_client_secret,
        refresh_token: settings.cmr_google_refresh_token,
      } : null,
    }),
  );
  const uploaded = JSON.parse(output.toString("utf8"));
  return {
    status: "uploaded",
    file_id: String(uploaded.id || ""),
    file_name: String(uploaded.name || filename),
    web_link: String(uploaded.webViewLink || uploaded.webContentLink || ""),
    mime_type: String(uploaded.mimeType || filePayload.type || ""),
    folder_id: countryFolderId,
    error: "",
  };
}

async function downloadFustDocumentFromDrive(documentInfo, settings) {
  const output = await runPythonBridge(
    ["drive-download-file"],
    JSON.stringify({
      file_id: documentInfo.file_id,
      oauth: settings.cmr_google_refresh_token ? {
        client_id: settings.cmr_google_client_id,
        client_secret: settings.cmr_google_client_secret,
        refresh_token: settings.cmr_google_refresh_token,
      } : null,
    }),
  );
  return output;
}

function contentDispositionFilename(filename) {
  return String(filename || "fust-document")
    .replace(/[\r\n"]/g, "_")
    .slice(0, 160) || "fust-document";
}

async function syncFustActionToSheets(action, settings) {
  if (!settings.spreadsheet_id) {
    return { ok: false, target_sheets: [], error: "Spreadsheet ID is not configured" };
  }

  const targetSheet = action.type === "OUT" ? settings.out_sheet_name : settings.in_sheet_name;
  if (!targetSheet) {
    return { ok: false, target_sheets: [], error: "Target sheet is not configured" };
  }

  await writeSheetRowToFirstEmpty(settings.spreadsheet_id, targetSheet, fustSheetRow(action));

  return {
    ok: true,
    target_sheets: [targetSheet],
    error: "",
    synced_at: new Date().toISOString(),
  };
}

function buildEmailMessage(action) {
  return [
    `Fust ${action.type} action`,
    "",
    `Action ID: ${action.id}`,
    `Date: ${action.action_date}`,
    `Day: ${action.day_name}`,
    `Week: ${action.week ?? ""}`,
    `Country: ${action.country}`,
    `Customer: ${action.customer_name}`,
    `Connect: ${action.connect_name || action.customer_code}`,
    `Remark: ${action.remark || "-"}`,
    "",
    `DC: ${action.metrics.dc}`,
    `CCTAG: ${action.metrics.cctag}`,
    `DCS: ${action.metrics.dcs}`,
    `DCO: ${action.metrics.dco}`,
    `PAL: ${action.metrics.pal}`,
    `VK: ${action.metrics.vk}`,
    "",
    `Created by: ${action.created_by}`,
    `Created at: ${action.created_at}`,
  ].join("\n");
}

async function sendFustActionEmail(action, settings) {
  const recipients = normalizeEmailRecipients(settings.email_recipients);
  if (!recipients.length) {
    return { ok: false, recipients: [], error: "No email recipients configured" };
  }

  await runPythonBridge(
    ["email-send"],
    JSON.stringify({
      recipients,
      subject: `Fust ${action.type} | ${action.country} | ${action.customer_name}`,
      body: buildEmailMessage(action),
      smtp: {
        host: settings.smtp_host,
        port: settings.smtp_port,
        username: settings.smtp_username,
        password: settings.smtp_password,
        from: settings.smtp_from,
        starttls: settings.smtp_starttls,
      },
    }),
  );

  return {
    ok: true,
    recipients,
    error: "",
    sent_at: new Date().toISOString(),
  };
}

function googleRunDetailsCachePath(folderId, accountName = "default") {
  return path.join(
    googleRunDetailsCacheDir,
    `${encodeURIComponent(accountName)}-${encodeURIComponent(folderId)}.json`,
  );
}

function isFreshTimestamp(value, ttlMinutes) {
  if (!ttlMinutes) {
    return false;
  }
  const parsed = Date.parse(String(value || ""));
  if (Number.isNaN(parsed)) {
    return false;
  }
  return Date.now() - parsed <= ttlMinutes * 60 * 1000;
}

async function readCachedGoogleRunDetails(run) {
  const accountName = String(run?.metadata?.drive_account || "default");
  const cachePath = googleRunDetailsCachePath(run.folder_id, accountName);
  const payload = await readJsonFile(cachePath, null);
  if (!payload || !isFreshTimestamp(payload.cached_at, googleRunDetailsCacheTtlMinutes)) {
    return null;
  }

  return {
    ...run,
    images: Array.isArray(payload.images) ? payload.images : [],
    qr_info: payload.qr_info || "No QR info found",
    qr_source: payload.qr_source || null,
  };
}

async function writeCachedGoogleRunDetails(run) {
  const accountName = String(run?.metadata?.drive_account || "default");
  const cachePath = googleRunDetailsCachePath(run.folder_id, accountName);
  await writeJsonFile(cachePath, {
    cached_at: new Date().toISOString(),
    images: Array.isArray(run.images) ? run.images : [],
    qr_info: run.qr_info || "No QR info found",
    qr_source: run.qr_source || null,
  });
}

async function hydrateGoogleRuns(runs) {
  const googleRuns = runs.filter((run) => String(run?.metadata?.source || "google_drive") === "google_drive");
  if (!googleRuns.length) {
    return new Map();
  }

  const cachedRuns = await Promise.all(googleRuns.map(readCachedGoogleRunDetails));
  const hydratedByFolderId = new Map();
  const missingRuns = [];

  for (let index = 0; index < googleRuns.length; index += 1) {
    const cachedRun = cachedRuns[index];
    if (cachedRun) {
      hydratedByFolderId.set(cachedRun.folder_id, cachedRun);
    } else {
      missingRuns.push(googleRuns[index]);
    }
  }

  if (missingRuns.length) {
    const output = await runPythonBridge(["details"], JSON.stringify(missingRuns));
    const fetchedRuns = JSON.parse(output.toString("utf8"));
    await Promise.all(fetchedRuns.map(writeCachedGoogleRunDetails));
    for (const run of fetchedRuns) {
      hydratedByFolderId.set(run.folder_id, run);
    }
  }

  return hydratedByFolderId;
}

async function isIndexedLocalImage(imagePath) {
  const normalizedImagePath = path.resolve(imagePath);
  const payload = await readRunData();

  for (const run of payload.runs) {
    if (String(run?.metadata?.source || "") !== "local_archive") {
      continue;
    }

    const runFolder = path.resolve(String(run.folder_id || ""));
    if (
      normalizedImagePath.startsWith(`${runFolder}${path.sep}`) &&
      imageExtensions.has(path.extname(normalizedImagePath).toLowerCase())
    ) {
      return true;
    }
  }

  return false;
}

function googleImageCachePath(fileId, accountName = "default") {
  return path.join(
    googleImageCacheDir,
    `${encodeURIComponent(accountName)}-${encodeURIComponent(fileId)}.bin`,
  );
}

async function readGoogleImage(fileId, accountName = "default", options = {}) {
  const { forceRefresh = false } = options;
  const cachePath = googleImageCachePath(fileId, accountName);
  if (!forceRefresh && existsSync(cachePath)) {
    return fs.readFile(cachePath);
  }

  const imageBytes = await runPythonBridge(["image", "--account", accountName, fileId]);
  await fs.mkdir(googleImageCacheDir, { recursive: true });
  await fs.writeFile(cachePath, imageBytes);
  return imageBytes;
}

async function listLocalRunDetails(run) {
  const source = String(run?.metadata?.source || "google_drive");
  if (source !== "local_archive") {
    return run;
  }

  const folderId = String(run.folder_id || "");
  let entries = [];
  try {
    entries = await fs.readdir(folderId, { withFileTypes: true });
  } catch {
    return { ...run, images: [], qr_info: "No QR info found", qr_source: null };
  }

  const files = entries
    .filter((entry) => entry.isFile())
    .map((entry) => path.join(folderId, entry.name));

  const images = [];
  for (const filePath of files) {
    const extension = path.extname(filePath).toLowerCase();
    if (!imageExtensions.has(extension)) {
      continue;
    }

    let stat = null;
    try {
      stat = await fs.stat(filePath);
    } catch {
      continue;
    }

    images.push({
      id: filePath,
      name: path.basename(filePath),
      mime_type: guessMimeType(filePath),
      web_view_link: null,
      size: stat.size,
    });
  }

  let qrInfo = "No QR info found";
  let qrSource = null;
  for (const filePath of files) {
    const name = path.basename(filePath).toLowerCase();
    if (name !== "qr.txt" && name !== "qr.json") {
      continue;
    }

    const content = await fs.readFile(filePath, "utf8").catch(() => "");
    qrSource = path.basename(filePath);
    if (name.endsWith(".json")) {
      try {
        qrInfo = JSON.stringify(JSON.parse(content), null, 2);
      } catch {
        qrInfo = content.trim() || "No QR info found";
      }
    } else {
      qrInfo = content.trim() || "No QR info found";
    }
    break;
  }

  if (!qrSource) {
    const qrName = files.map((filePath) => path.basename(filePath)).find((name) => name.toLowerCase().includes("qr"));
    if (qrName) {
      qrInfo = qrName;
      qrSource = "filename";
    }
  }

  return { ...run, images, qr_info: qrInfo, qr_source: qrSource };
}

function groupByCustomer(runs) {
  const groups = new Map();
  for (const run of runs) {
    const customer = run.customer_code || "Unknown";
    if (!groups.has(customer)) {
      groups.set(customer, []);
    }
    groups.get(customer).push(run);
  }

  return [...groups.entries()]
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([customer_code, customerRuns]) => ({
      customer_code,
      runs: customerRuns.sort((left, right) => {
        const carrierCompare = String(left.carrier || "").localeCompare(String(right.carrier || ""));
        if (carrierCompare !== 0) {
          return carrierCompare;
        }
        return String(left.run_id || "").localeCompare(String(right.run_id || ""));
      }),
    }));
}

function startSync(mode, selectedDate) {
  if (!existsSync(syncScriptPath) || !existsSync(syncWorkerPath)) {
    return false;
  }

  const args = [syncWorkerPath, "--mode", mode];
  if (selectedDate) {
    args.push("--date", selectedDate);
  }

  const child = spawn(process.execPath, args, {
    cwd: repoRoot,
    detached: true,
    stdio: "ignore",
    windowsHide: true,
  });
  child.unref();
  return true;
}

async function maybeStartAutoSync(payload, activeDate) {
  if (!autoSyncOnVisit || !existsSync(syncScriptPath)) {
    return null;
  }
  if (await isSyncRunning()) {
    return null;
  }

  const mode = "rebuild";
  const selectedDate = null;
  const throttleKey = `${mode}:${selectedDate || "all"}`;
  const now = Date.now();
  const previousStart = autoSyncStartedAt.get(throttleKey) || 0;

  if (now - previousStart < autoSyncThrottleMs) {
    return null;
  }

  if (startSync(mode, selectedDate)) {
    autoSyncStartedAt.set(throttleKey, now);
    return {
      mode,
      date: selectedDate,
      started_at: new Date().toISOString(),
    };
  }

  return null;
}

async function pruneGoogleImageCache(keepCacheNames) {
  if (!keepCacheNames.size || !existsSync(googleImageCacheDir)) {
    return;
  }

  const entries = await fs.readdir(googleImageCacheDir, { withFileTypes: true });
  await Promise.all(entries.map(async (entry) => {
    if (!entry.isFile() || keepCacheNames.has(entry.name)) {
      return;
    }
    try {
      await fs.unlink(path.join(googleImageCacheDir, entry.name));
    } catch {
      // Cache pruning is best effort. A failed delete should not block the app.
    }
  }));
}

async function preloadRecentGoogleImages(payload) {
  if (!recentPreloadDays || recentPreloadMaxImages === 0) {
    return;
  }

  const dates = [...new Set(payload.runs.map((run) => run.run_date).filter(Boolean))].sort().reverse();
  const selectedDates = new Set(dates.slice(0, recentPreloadDays));
  if (!selectedDates.size) {
    return;
  }

  const recentRuns = payload.runs.filter((run) => selectedDates.has(run.run_date));
  if (!recentRuns.length) {
    return;
  }

  const hydratedByFolderId = await hydrateGoogleRuns(recentRuns);
  const keepCacheNames = new Set();
  let remaining = recentPreloadMaxImages < 0 ? Infinity : recentPreloadMaxImages;

  for (const run of recentRuns) {
    const hydrated = hydratedByFolderId.get(run.folder_id);
    if (!hydrated || !Array.isArray(hydrated.images)) {
      continue;
    }

    const accountName = String(hydrated?.metadata?.drive_account || "default");
    for (const image of hydrated.images) {
      if (!image?.id) {
        continue;
      }
      keepCacheNames.add(path.basename(googleImageCachePath(image.id, accountName)));
      if (remaining <= 0) {
        continue;
      }
      try {
        await readGoogleImage(image.id, accountName);
        remaining -= 1;
      } catch {
        // Ignore individual preload failures so the dashboard can still load normally.
      }
    }
  }

  await pruneGoogleImageCache(keepCacheNames);
}

function maybeStartRecentPreload(payload) {
  if (!payload?.generated_at || !recentPreloadDays || recentPreloadMaxImages === 0) {
    return;
  }

  const cacheKey = String(payload.generated_at);
  const previousStart = recentPreloadStartedAt.get(cacheKey) || 0;
  if (previousStart) {
    return;
  }

  recentPreloadStartedAt.set(cacheKey, Date.now());
  void preloadRecentGoogleImages(payload).catch(() => {
    // Best-effort warmup only.
  });
}

function publicBaseUrl(req) {
  const proto = String(req.headers["x-forwarded-proto"] || "http").split(",")[0].trim();
  const host = String(req.headers["x-forwarded-host"] || req.headers.host || "").split(",")[0].trim();
  return `${proto}://${host}`;
}

function cmrGoogleRedirectUri(req) {
  return `${publicBaseUrl(req)}/api/fust/google/callback`;
}

function cmrGoogleAuthUrl(settings, req) {
  const params = new URLSearchParams({
    client_id: settings.cmr_google_client_id,
    redirect_uri: cmrGoogleRedirectUri(req),
    response_type: "code",
    scope: "https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/userinfo.email openid",
    access_type: "offline",
    prompt: "consent",
  });
  return `https://accounts.google.com/o/oauth2/v2/auth?${params.toString()}`;
}

async function exchangeGoogleAuthCode(settings, req, code) {
  const response = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      code,
      client_id: settings.cmr_google_client_id,
      client_secret: settings.cmr_google_client_secret,
      redirect_uri: cmrGoogleRedirectUri(req),
      grant_type: "authorization_code",
    }),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(payload.error_description || payload.error || `Google token exchange failed with ${response.status}`);
  }
  return payload;
}

async function loadGoogleUserEmail(accessToken) {
  if (!accessToken) {
    return "";
  }
  const response = await fetch("https://www.googleapis.com/oauth2/v2/userinfo", {
    headers: { authorization: `Bearer ${accessToken}` },
  });
  const payload = await response.json().catch(() => ({}));
  return String(payload.email || "");
}

async function handleApi(req, res, url) {
  if (url.pathname === "/api/auth/me") {
    const users = await readUsers();
    const user = await getRequestUser(req);
    sendJson(res, 200, {
      user,
      setup_required: users.length === 0,
    });
    return;
  }

  if (url.pathname === "/api/auth/setup" && req.method === "POST") {
    const users = await readUsers();
    if (users.length > 0) {
      sendJson(res, 409, { error: "Setup is already complete" });
      return;
    }

    const body = await readRequestJson(req);
    const username = String(body.username || "").trim();
    const password = String(body.password || "");
    if (!username || password.length < 6) {
      sendJson(res, 400, { error: "Username is required and password must be at least 6 characters" });
      return;
    }

    const user = {
      username,
      role: "admin",
      permissions: [...roleDefaultPermissions.admin],
      password_hash: hashPassword(password),
      created_at: new Date().toISOString(),
    };
    await writeUsers([user]);
    createSession(res, user);
    sendJson(res, 201, { user: publicUser(user) });
    return;
  }

  if (url.pathname === "/api/auth/login" && req.method === "POST") {
    const body = await readRequestJson(req);
    const username = String(body.username || "").trim();
    const password = String(body.password || "");
    const users = await readUsers();
    const user = users.find((item) => item.username.toLowerCase() === username.toLowerCase());
    if (!user || !verifyPassword(password, user.password_hash)) {
      sendJson(res, 401, { error: "Invalid username or password" });
      return;
    }

    createSession(res, user);
    sendJson(res, 200, { user: publicUser(user) });
    return;
  }

  if (url.pathname === "/api/auth/logout" && req.method === "POST") {
    destroySession(req, res);
    sendJson(res, 200, { ok: true });
    return;
  }

  const requestUser = await getRequestUser(req);
  if (!requestUser) {
    sendUnauthorized(res);
    return;
  }

  if (url.pathname === "/api/clock/employees") {
    if (!requirePermission(res, requestUser, PERMISSIONS.CLOCK_VIEW)) {
      return;
    }
    const settings = await readFustSettings();
    try {
      const rows = await loadSheetRows(settings.clock_spreadsheet_id, settings.clock_employee_sheet_name);
      const parsed = buildClockEmployeesFromSheetRows(rows);
      sendJson(res, 200, {
        employees: parsed.employees,
        headers: parsed.headers,
        raw_row_count: parsed.raw_row_count,
        sheet_name: settings.clock_employee_sheet_name,
      });
    } catch (error) {
      sendJson(res, 500, { error: error instanceof Error ? error.message : String(error), employees: [] });
    }
    return;
  }

  if (url.pathname === "/api/clock/records/export" && req.method === "GET") {
    if (!requirePermission(res, requestUser, PERMISSIONS.CLOCK_VIEW)) {
      return;
    }
    const fromDate = String(url.searchParams.get("from") || url.searchParams.get("date") || localDateIso()).slice(0, 10);
    const toDate = String(url.searchParams.get("to") || fromDate).slice(0, 10);
    const [startDate, endDate] = fromDate <= toDate ? [fromDate, toDate] : [toDate, fromDate];
    const records = addClockWorkedDurations(await readClockRecords())
      .filter((record) => record.action_date >= startDate && record.action_date <= endDate)
      .sort((left, right) => `${left.action_date}T${left.action_time}`.localeCompare(`${right.action_date}T${right.action_time}`));
    const filenameDate = startDate === endDate ? startDate : `${startDate}-to-${endDate}`;
    res.writeHead(200, {
      "content-type": "text/tab-separated-values; charset=utf-8",
      "content-disposition": `attachment; filename="clock-times-${filenameDate}.tsv"`,
      "cache-control": "no-store",
    });
    res.end(clockExportText(records));
    return;
  }

  if (url.pathname === "/api/clock/records/import-backup" && req.method === "POST") {
    if (!requirePermission(res, requestUser, PERMISSIONS.CLOCK_MANAGE)) {
      return;
    }
    const body = await readRequestJson(req);
    const date = String(body.date || localDateIso()).slice(0, 10);
    const settings = await readFustSettings();
    if (!settings.clock_spreadsheet_id || !settings.clock_records_sheet_name) {
      sendJson(res, 400, { error: "Clock spreadsheet ID and backup tab must be configured" });
      return;
    }

    const rows = await loadSheetRows(settings.clock_spreadsheet_id, settings.clock_records_sheet_name);
    const importedRecords = parseClockBackupRows(rows, date, requestUser.username);
    const existingRecords = await readClockRecords();
    const keptRecords = existingRecords.filter((record) => record.action_date !== date);
    const nextRecords = [...keptRecords, ...importedRecords];
    await writeClockRecords(nextRecords);

    sendJson(res, 200, {
      date,
      imported_count: importedRecords.length,
      records: filterClockRecords(nextRecords, date),
      sessions: clockSessionRows(importedRecords).map((session) => ({ in_record: session.inRecord, out_record: session.outRecord, row: session.row })),
    });
    return;
  }

  if (url.pathname.startsWith("/api/clock/records/") && req.method !== "GET") {
    const recordId = decodeURIComponent(url.pathname.slice("/api/clock/records/".length));
    if (!requirePermission(res, requestUser, PERMISSIONS.CLOCK_MANAGE)) {
      return;
    }
    const records = await readClockRecords();
    const recordIndex = records.findIndex((record) => record.id === recordId);
    if (recordIndex < 0) {
      sendJson(res, 404, { error: "Clock record not found" });
      return;
    }

    if (req.method === "DELETE") {
      records.splice(recordIndex, 1);
      await writeClockRecords(records);
      sendJson(res, 200, { ok: true, deleted_record_id: recordId });
      return;
    }

    if (req.method === "PATCH") {
      const body = await readRequestJson(req);
      const current = records[recordIndex];
      const updated = normalizeClockRecord({
        ...current,
        action_date: body.action_date ?? current.action_date,
        action_time: body.action_time ?? current.action_time,
        timestamp: `${body.action_date ?? current.action_date}T${body.action_time ?? current.action_time}`,
        tbnr: body.tbnr ?? current.tbnr,
        name: body.name ?? current.name,
        employee_type: body.employee_type ?? current.employee_type,
        direction: body.direction ?? current.direction,
        source: current.source,
        updated_by: requestUser.username,
        updated_at: new Date().toISOString(),
      });
      records[recordIndex] = updated;
      await writeClockRecords(records);
      sendJson(res, 200, { record: updated, records: filterClockRecords(records, updated.action_date), sessions: clockSessionRows(records.filter((item) => item.action_date === updated.action_date)).map((session) => ({ in_record: session.inRecord, out_record: session.outRecord, row: session.row })) });
      return;
    }
  }

  if (url.pathname === "/api/clock/records") {
    if (!requirePermission(res, requestUser, PERMISSIONS.CLOCK_VIEW)) {
      return;
    }

    if (req.method === "GET") {
      const date = String(url.searchParams.get("date") || localDateIso()).slice(0, 10);
      const records = await readClockRecords();
      sendJson(res, 200, { date, records: filterClockRecords(records, date), sessions: clockSessionRows(records.filter((record) => record.action_date === date)).map((session) => ({ in_record: session.inRecord, out_record: session.outRecord, row: session.row })) });
      return;
    }

    if (req.method === "POST") {
      if (!requirePermission(res, requestUser, PERMISSIONS.CLOCK_MANAGE)) {
        return;
      }
      const body = await readRequestJson(req);
      const employee = normalizeClockEmployee(body.employee || {
        tbnr: body.tbnr,
        name: body.name,
        type: body.employee_type,
      });
      if (!employee.tbnr || !employee.name) {
        sendJson(res, 400, { error: "Choose a valid employee" });
        return;
      }
      const actionDate = String(body.action_date || localDateIso()).slice(0, 10);
      const actionTime = String(body.action_time || new Date().toLocaleTimeString("nl-NL", { hour12: false })).slice(0, 8);
      const direction = String(body.direction || "IN").toUpperCase() === "OUT" ? "OUT" : "IN";
      const record = createClockRecord(employee, direction, actionDate, actionTime, "manual", requestUser.username);
      const records = await readClockRecords();
      records.push(record);
      await writeClockRecords(records);
      const settings = await readFustSettings();
      try {
        record.sheet_sync = await syncClockRecordToSheets(record, settings, records);
      } catch (error) {
        record.sheet_sync = { ok: false, target_sheets: [], error: error instanceof Error ? error.message : String(error) };
      }
      const savedRecords = await readClockRecords();
      const savedIndex = savedRecords.findIndex((item) => item.id === record.id);
      if (savedIndex >= 0) {
        savedRecords[savedIndex] = record;
        await writeClockRecords(savedRecords);
      }
      sendJson(res, 201, { record, records: filterClockRecords(savedRecords, actionDate), sessions: clockSessionRows(savedRecords.filter((item) => item.action_date === actionDate)).map((session) => ({ in_record: session.inRecord, out_record: session.outRecord, row: session.row })) });
      return;
    }
  }

  if (url.pathname === "/api/clock/scan" && req.method === "POST") {
    if (!requirePermission(res, requestUser, PERMISSIONS.CLOCK_VIEW)) {
      return;
    }
    const body = await readRequestJson(req);
    const code = String(body.code || "").trim().toUpperCase();
    if (!code) {
      sendJson(res, 400, { error: "Scan a badge code first" });
      return;
    }
    const settings = await readFustSettings();
    const rows = await loadSheetRows(settings.clock_spreadsheet_id, settings.clock_employee_sheet_name);
    const employee = buildClockEmployeesFromSheetRows(rows).employees.find((item) => item.tbnr === code);
    if (!employee) {
      sendJson(res, 404, { error: `${code} is not in the employee sheet` });
      return;
    }
    const now = new Date();
    const actionDate = String(body.action_date || localDateIso()).slice(0, 10);
    const actionTime = String(body.action_time || now.toLocaleTimeString("nl-NL", { hour12: false })).slice(0, 8);
    const records = await readClockRecords();
    const direction = nextClockDirection(records, code, actionDate);
    const record = createClockRecord(employee, direction, actionDate, actionTime, "scanner", requestUser.username);
    records.push(record);
    await writeClockRecords(records);
    try {
      record.sheet_sync = await syncClockRecordToSheets(record, settings, records);
    } catch (error) {
      record.sheet_sync = { ok: false, target_sheets: [], error: error instanceof Error ? error.message : String(error) };
    }
    const savedRecords = await readClockRecords();
    const savedIndex = savedRecords.findIndex((item) => item.id === record.id);
    if (savedIndex >= 0) {
      savedRecords[savedIndex] = record;
      await writeClockRecords(savedRecords);
    }
    sendJson(res, 201, { record, records: filterClockRecords(savedRecords, actionDate), sessions: clockSessionRows(savedRecords.filter((item) => item.action_date === actionDate)).map((session) => ({ in_record: session.inRecord, out_record: session.outRecord, row: session.row })) });
    return;
  }

  if (url.pathname === "/api/fust/settings") {
    if (!requirePermission(res, requestUser, PERMISSIONS.SETTINGS_MANAGE)) {
      return;
    }

    if (req.method === "GET") {
      sendJson(res, 200, { settings: await readFustSettings() });
      return;
    }

    if (req.method === "PATCH") {
      const body = await readRequestJson(req);
      const currentSettings = await readFustSettings();
      const nextSettings = normalizeFustSettings({
        ...currentSettings,
        ...body,
      });
      await writeFustSettings(nextSettings);
      sendJson(res, 200, { settings: nextSettings });
      return;
    }
  }

  if (url.pathname === "/api/fust/google/auth-url") {
    if (!requirePermission(res, requestUser, PERMISSIONS.SETTINGS_MANAGE)) {
      return;
    }
    const settings = await readFustSettings();
    if (!settings.cmr_google_client_id || !settings.cmr_google_client_secret) {
      sendJson(res, 400, { error: "Set Google OAuth client ID and secret first" });
      return;
    }
    sendJson(res, 200, { auth_url: cmrGoogleAuthUrl(settings, req), redirect_uri: cmrGoogleRedirectUri(req) });
    return;
  }

  if (url.pathname === "/api/fust/google/callback") {
    if (!requirePermission(res, requestUser, PERMISSIONS.SETTINGS_MANAGE)) {
      return;
    }
    const code = String(url.searchParams.get("code") || "");
    if (!code) {
      sendText(res, 400, "Missing Google authorization code");
      return;
    }
    try {
      const settings = await readFustSettings();
      const tokenPayload = await exchangeGoogleAuthCode(settings, req, code);
      if (!tokenPayload.refresh_token) {
        sendText(res, 400, "Google did not return a refresh token. Try Connect Google Drive again and approve offline access.");
        return;
      }
      const connectedEmail = await loadGoogleUserEmail(tokenPayload.access_token);
      await writeFustSettings({
        ...settings,
        cmr_google_refresh_token: tokenPayload.refresh_token,
        cmr_google_connected_email: connectedEmail,
      });
      res.writeHead(200, { "content-type": "text/html; charset=utf-8" });
      res.end("<p>Google Drive connected. You can close this tab and return to SnappySjaak Settings.</p>");
    } catch (error) {
      sendText(res, 500, error instanceof Error ? error.message : String(error));
    }
    return;
  }

  if (url.pathname === "/api/fust/backups") {
    if (!requirePermission(res, requestUser, PERMISSIONS.SETTINGS_MANAGE)) {
      return;
    }

    if (req.method === "GET") {
      sendJson(res, 200, { backups: await listFustBackups() });
      return;
    }

    if (req.method === "POST") {
      const backup = await createFustBackupSnapshot(requestUser.username);
      sendJson(res, 201, { backup, backups: await listFustBackups() });
      return;
    }
  }

  if (url.pathname.startsWith("/api/fust/backups/") && req.method === "GET") {
    if (!requirePermission(res, requestUser, PERMISSIONS.SETTINGS_MANAGE)) {
      return;
    }

    const filename = decodeURIComponent(url.pathname.slice("/api/fust/backups/".length));
    const resolvedPath = path.resolve(fustBackupDir, filename);
    if (!resolvedPath.startsWith(path.resolve(fustBackupDir))) {
      sendText(res, 403, "Forbidden");
      return;
    }
    if (!existsSync(resolvedPath)) {
      sendText(res, 404, "Backup not found");
      return;
    }

    res.writeHead(200, {
      "content-type": "application/json; charset=utf-8",
      "content-disposition": `attachment; filename="${path.basename(resolvedPath)}"`,
      "cache-control": "no-store",
    });
    createReadStream(resolvedPath).pipe(res);
    return;
  }

  if (url.pathname === "/api/fust/meta") {
    if (!requirePermission(res, requestUser, PERMISSIONS.FUST_VIEW)) {
      return;
    }

    const settings = await readFustSettings();
    let records = [];
    let headers = [];
    let rawRowCount = 0;
    let source = "local";
    let error = "";

    try {
      const rows = await loadFustSheetRows(settings);
      const parsed = buildFustMetaFromSheetRows(rows);
      records = parsed.records;
      headers = parsed.headers;
      rawRowCount = parsed.raw_row_count;
      source = rows.length ? "spreadsheet" : "local";
    } catch (sheetError) {
      error = sheetError instanceof Error ? sheetError.message : String(sheetError);
    }

    const countries = [...new Set(records.map((record) => record.country))].sort((left, right) => left.localeCompare(right));
    sendJson(res, 200, {
      settings,
      countries,
      records,
      headers,
      raw_row_count: rawRowCount,
      sample_records: records.slice(0, 8),
      source,
      error,
    });
    return;
  }

  if (url.pathname === "/api/fust/connection-test") {
    if (!requirePermission(res, requestUser, PERMISSIONS.SETTINGS_MANAGE)) {
      return;
    }

    const settings = await readFustSettings();
    let account = { client_email: "", project_id: "" };
    let read_ok = false;
    let row_count = 0;
    let headers = [];
    let error = "";

    try {
      account = await loadServiceAccountInfo();
    } catch (accountError) {
      error = accountError instanceof Error ? accountError.message : String(accountError);
    }

    if (!error) {
      try {
        const rows = await loadFustSheetRows(settings);
        read_ok = true;
        row_count = rows.length;
        headers = Array.isArray(rows[0]) ? rows[0].map((value) => String(value || "").trim()) : [];
      } catch (sheetError) {
        error = sheetError instanceof Error ? sheetError.message : String(sheetError);
      }
    }

    sendJson(res, 200, {
      account,
      spreadsheet_id: settings.spreadsheet_id,
      sheet_name: settings.data_sheet_name,
      read_ok,
      row_count,
      headers,
      error,
    });
    return;
  }

  if (url.pathname.startsWith("/api/fust/actions/") && req.method === "GET") {
    const parts = url.pathname.split("/").filter(Boolean);
    const actionId = decodeURIComponent(parts[3] || "");
    const routeKind = parts[4] || "";
    const documentKind = parts[5] || "";
    if (routeKind === "document") {
      if (!requirePermission(res, requestUser, PERMISSIONS.FUST_VIEW)) {
        return;
      }
      if (!["cmr", "fustbon"].includes(documentKind)) {
        sendText(res, 404, "Document not found");
        return;
      }
      const actions = await readFustActions();
      const action = actions.find((item) => item.id === actionId);
      const documentInfo = normalizeCmrInfo(action?.[documentKind]);
      if (!action || documentInfo.status !== "uploaded" || !documentInfo.file_id) {
        sendText(res, 404, "Document not found");
        return;
      }
      try {
        const settings = await readFustSettings();
        const fileBuffer = await downloadFustDocumentFromDrive(documentInfo, settings);
        const fileName = contentDispositionFilename(documentInfo.file_name || `${documentKind}-${actionId}`);
        res.writeHead(200, {
          "content-type": documentInfo.mime_type || guessMimeType(fileName),
          "content-disposition": `inline; filename="${fileName}"`,
          "cache-control": "private, no-store",
        });
        res.end(fileBuffer);
      } catch (error) {
        sendText(res, 500, error instanceof Error ? error.message : String(error));
      }
      return;
    }
  }

  if (url.pathname === "/api/fust/actions") {
    if (!requirePermission(res, requestUser, PERMISSIONS.FUST_VIEW)) {
      return;
    }

    const settings = await readFustSettings();
    const localActions = await readFustActions();
    let inSheetActions = [];
    let outSheetActions = [];
    let sheetActions = [];
    const sourceDebug = {
      local: {
        action_count: localActions.length,
      },
      in_sheet: {
        sheet_name: settings.in_sheet_name,
        row_count: 0,
        action_count: 0,
        error: "",
      },
      out_sheet: {
        sheet_name: settings.out_sheet_name,
        row_count: 0,
        action_count: 0,
        error: "",
      },
      dashboard_sheet: {
        sheet_name: settings.dashboard_sheet_name,
        row_count: 0,
        action_count: 0,
        error: "",
      },
    };
    try {
      const retourRows = await loadSheetRows(settings.spreadsheet_id, settings.in_sheet_name);
      sourceDebug.in_sheet.row_count = retourRows.length;
      inSheetActions = parseRegistrySheetRows(retourRows, "IN");
      sourceDebug.in_sheet.action_count = inSheetActions.length;
    } catch (error) {
      inSheetActions = [];
      sourceDebug.in_sheet.error = error instanceof Error ? error.message : String(error || "Unknown error");
    }

    try {
      const uitgaandRows = await loadSheetRows(settings.spreadsheet_id, settings.out_sheet_name);
      sourceDebug.out_sheet.row_count = uitgaandRows.length;
      outSheetActions = parseRegistrySheetRows(uitgaandRows, "OUT");
      sourceDebug.out_sheet.action_count = outSheetActions.length;
    } catch (error) {
      outSheetActions = [];
      sourceDebug.out_sheet.error = error instanceof Error ? error.message : String(error || "Unknown error");
    }

    try {
      const dashboardRows = await loadSheetRows(settings.spreadsheet_id, settings.dashboard_sheet_name);
      sourceDebug.dashboard_sheet.row_count = dashboardRows.length;
      sheetActions = parseDashboardSheetRows(dashboardRows);
      sourceDebug.dashboard_sheet.action_count = sheetActions.length;
    } catch (error) {
      sheetActions = [];
      sourceDebug.dashboard_sheet.error = error instanceof Error ? error.message : String(error || "Unknown error");
    }

    const dedupedActions = new Map();
    for (const action of [...inSheetActions, ...outSheetActions, ...sheetActions, ...localActions]) {
      dedupedActions.set(buildActionSignature(action), action);
    }

    const actions = [...dedupedActions.values()];
    const country = String(url.searchParams.get("country") || "").trim();
    const customer = String(url.searchParams.get("customer_name") || "").trim().toLowerCase();
    const type = String(url.searchParams.get("type") || "").trim().toUpperCase();
    const filteredActions = actions
      .filter((action) => !country || action.country === country)
      .filter((action) => !customer || action.customer_name.toLowerCase().includes(customer))
      .filter((action) => !type || action.type === type);

    sendJson(res, 200, {
      actions: filteredActions.sort((left, right) => {
        const rightDate = String(right.created_at || right.action_date || "");
        const leftDate = String(left.created_at || left.action_date || "");
        return rightDate.localeCompare(leftDate);
      }),
      overview: buildOverview(filteredActions),
      source_debug: {
        ...sourceDebug,
        merged_action_count: actions.length,
        filtered_action_count: filteredActions.length,
      },
    });
    return;
  }

  if (url.pathname.startsWith("/api/fust/actions/") && req.method === "PATCH") {
    const parts = url.pathname.split("/").filter(Boolean);
    const actionId = decodeURIComponent(parts[3] || "");
    const documentAction = parts[4] || "";
    const documentConfig = {
      "cmr-upload": { field: "cmr", type: "OUT", label: "CMR", mode: "upload" },
      "cmr-skip": { field: "cmr", type: "OUT", label: "CMR", mode: "skip" },
      "fustbon-upload": { field: "fustbon", type: "IN", label: "Fustbon", mode: "upload" },
      "fustbon-skip": { field: "fustbon", type: "IN", label: "Fustbon", mode: "skip" },
    }[documentAction];
    if (!documentConfig) {
      sendJson(res, 404, { error: "Unknown action update" });
      return;
    }

    const actions = await readFustActions();
    const actionIndex = actions.findIndex((item) => item.id === actionId);
    if (actionIndex < 0) {
      sendJson(res, 404, { error: "Fust action not found" });
      return;
    }

    const action = actions[actionIndex];
    if (action.type !== documentConfig.type) {
      sendJson(res, 400, { error: `${documentConfig.label} files can only be attached to ${documentConfig.type} actions` });
      return;
    }
    const requiredPermission = action.type === "OUT" ? PERMISSIONS.FUST_OUT : PERMISSIONS.FUST_IN;
    if (!requirePermission(res, requestUser, requiredPermission)) {
      return;
    }

    if (documentConfig.mode === "skip") {
      action[documentConfig.field] = normalizeCmrInfo({
        status: "skipped",
        uploaded_at: new Date().toISOString(),
        uploaded_by: requestUser.username,
      });
      actions[actionIndex] = action;
      await writeFustActions(actions);
      sendJson(res, 200, { action });
      return;
    }

    const body = await readRequestJson(req, 18 * 1024 * 1024);
    const filePayload = body?.file || {};
    if (!filePayload.content_base64 || !filePayload.name) {
      sendJson(res, 400, { error: `Choose a ${documentConfig.label} file first` });
      return;
    }

    const settings = await readFustSettings();
    try {
      action[documentConfig.field] = normalizeCmrInfo({
        ...(await uploadFustDocumentToDrive(action, settings, filePayload, documentConfig.field)),
        uploaded_at: new Date().toISOString(),
        uploaded_by: requestUser.username,
      });
    } catch (documentError) {
      action[documentConfig.field] = normalizeCmrInfo({
        status: "failed",
        error: documentError instanceof Error ? documentError.message : String(documentError),
        uploaded_at: new Date().toISOString(),
        uploaded_by: requestUser.username,
      });
      actions[actionIndex] = action;
      await writeFustActions(actions);
      sendJson(res, 500, { error: action[documentConfig.field].error, action });
      return;
    }

    actions[actionIndex] = action;
    await writeFustActions(actions);
    sendJson(res, 200, { action });
    return;
  }

  if (url.pathname.startsWith("/api/fust/actions/") && req.method === "POST") {
    const parts = url.pathname.split("/").filter(Boolean);
    const actionId = decodeURIComponent(parts[3] || "");
    const retryKind = parts[4] || "";
    const actions = await readFustActions();
    const actionIndex = actions.findIndex((item) => item.id === actionId);
    if (actionIndex < 0) {
      sendJson(res, 404, { error: "Fust action not found" });
      return;
    }

    const action = actions[actionIndex];
    const settings = await readFustSettings();

    if (retryKind === "retry-sheet") {
      const requiredPermission = action.type === "OUT" ? PERMISSIONS.FUST_OUT : PERMISSIONS.FUST_IN;
      if (!requirePermission(res, requestUser, requiredPermission)) {
        return;
      }

      try {
        action.sheet_sync = await syncFustActionToSheets(action, settings);
      } catch (sheetError) {
        action.sheet_sync = {
          ok: false,
          target_sheets: [],
          error: sheetError instanceof Error ? sheetError.message : String(sheetError),
        };
      }

      actions[actionIndex] = action;
      await writeFustActions(actions);
      sendJson(res, 200, { action });
      return;
    }

    if (retryKind === "retry-email") {
      const requiredPermission = action.type === "OUT" ? PERMISSIONS.FUST_OUT : PERMISSIONS.FUST_IN;
      if (!requirePermission(res, requestUser, requiredPermission)) {
        return;
      }

      try {
        action.email_sync = await sendFustActionEmail(action, settings);
      } catch (emailError) {
        action.email_sync = {
          ok: false,
          recipients: normalizeEmailRecipients(settings.email_recipients),
          error: emailError instanceof Error ? emailError.message : String(emailError),
        };
      }

      actions[actionIndex] = action;
      await writeFustActions(actions);
      sendJson(res, 200, { action });
      return;
    }

    sendJson(res, 404, { error: "Unknown retry action" });
    return;
  }

  if (url.pathname.startsWith("/api/fust/actions/") && req.method === "PUT") {
    const parts = url.pathname.split("/").filter(Boolean);
    const actionId = decodeURIComponent(parts[3] || "");
    const actions = await readFustActions();
    const actionIndex = actions.findIndex((item) => item.id === actionId);
    if (actionIndex < 0) {
      sendJson(res, 404, { error: "Fust action not found" });
      return;
    }

    const existingAction = actions[actionIndex];
    const type = String(existingAction.type || "").trim().toUpperCase() === "OUT" ? "OUT" : "IN";
    const requiredPermission = type === "OUT" ? PERMISSIONS.FUST_OUT : PERMISSIONS.FUST_IN;
    if (!requirePermission(res, requestUser, requiredPermission)) {
      return;
    }

    const body = await readRequestJson(req);
    const actionDate = String(body.action_date || existingAction.action_date || localDateIso()).trim();
    const updatedAction = normalizeFustAction({
      ...existingAction,
      type: String(body.type || existingAction.type || "IN").trim().toUpperCase() === "OUT" ? "OUT" : "IN",
      action_date: actionDate,
      week: weekNumberForDate(actionDate),
      day_name: weekdayNameForDate(actionDate),
      country: body.country,
      customer_name: body.customer_name,
      customer_code: body.customer_code,
      connect_name: body.connect_name,
      remark: body.remark,
      metrics: body.metrics,
      sheet_sync: {
        ...(existingAction.sheet_sync || {}),
        ok: false,
        error: "Edited locally",
      },
      email_sync: {
        ...(existingAction.email_sync || {}),
        ok: false,
        error: "Edited locally",
      },
    });

    if (!updatedAction.country || !updatedAction.customer_name || !updatedAction.connect_name) {
      sendJson(res, 400, { error: "Country, customer, and connect are required" });
      return;
    }

    const newRequiredPermission = updatedAction.type === "OUT" ? PERMISSIONS.FUST_OUT : PERMISSIONS.FUST_IN;
    if (!requirePermission(res, requestUser, newRequiredPermission)) {
      return;
    }

    actions[actionIndex] = updatedAction;
    await writeFustActions(actions);
    sendJson(res, 200, { action: updatedAction });
    return;
  }

  if (url.pathname.startsWith("/api/fust/actions/") && req.method === "DELETE") {
    const parts = url.pathname.split("/").filter(Boolean);
    const actionId = decodeURIComponent(parts[3] || "");
    const actions = await readFustActions();
    const actionIndex = actions.findIndex((item) => item.id === actionId);
    if (actionIndex < 0) {
      sendJson(res, 404, { error: "Fust action not found" });
      return;
    }

    const action = actions[actionIndex];
    const requiredPermission = action.type === "OUT" ? PERMISSIONS.FUST_OUT : PERMISSIONS.FUST_IN;
    if (!requirePermission(res, requestUser, requiredPermission)) {
      return;
    }

    actions.splice(actionIndex, 1);
    await writeFustActions(actions);
    sendJson(res, 200, { ok: true, deleted_action_id: actionId });
    return;
  }

  if (url.pathname === "/api/fust/submit" && req.method === "POST") {
    const body = await readRequestJson(req);
    const type = String(body.type || "").trim().toUpperCase() === "OUT" ? "OUT" : "IN";
    const requiredPermission = type === "OUT" ? PERMISSIONS.FUST_OUT : PERMISSIONS.FUST_IN;
    if (!requirePermission(res, requestUser, requiredPermission)) {
      return;
    }

    const actionDate = String(body.action_date || localDateIso()).trim();
    const action = normalizeFustAction({
      id: crypto.randomUUID(),
      type,
      action_date: actionDate,
      week: weekNumberForDate(actionDate),
      day_name: weekdayNameForDate(actionDate),
      country: body.country,
      customer_name: body.customer_name,
      customer_code: body.customer_code,
      connect_name: body.connect_name,
      remark: body.remark,
      metrics: body.metrics,
      created_by: requestUser.username,
      created_at: new Date().toISOString(),
      sheet_sync: { ok: false, target_sheets: [], error: "Pending" },
      email_sync: { ok: false, recipients: [], error: "Pending" },
    });

    if (!action.country || !action.customer_name || !action.connect_name) {
      sendJson(res, 400, { error: "Country, customer, and connect are required" });
      return;
    }

    const actions = await readFustActions();
    actions.push(action);
    await writeFustActions(actions);

    const settings = await readFustSettings();
    try {
      action.sheet_sync = await syncFustActionToSheets(action, settings);
    } catch (sheetError) {
      action.sheet_sync = {
        ok: false,
        target_sheets: [],
        error: sheetError instanceof Error ? sheetError.message : String(sheetError),
      };
    }

    try {
      action.email_sync = await sendFustActionEmail(action, settings);
    } catch (emailError) {
      action.email_sync = {
        ok: false,
        recipients: normalizeEmailRecipients(settings.email_recipients),
        error: emailError instanceof Error ? emailError.message : String(emailError),
      };
    }

    const savedActions = await readFustActions();
    const actionIndex = savedActions.findIndex((item) => item.id === action.id);
    if (actionIndex >= 0) {
      savedActions[actionIndex] = action;
      await writeFustActions(savedActions);
    }

    sendJson(res, 201, { action });
    return;
  }

  if (url.pathname === "/api/users") {
    if (requestUser.role !== "admin") {
      sendForbidden(res);
      return;
    }

    if (req.method === "GET") {
      const users = await readUsers();
      sendJson(res, 200, { users: users.map(publicUser) });
      return;
    }

    if (req.method === "POST") {
      const body = await readRequestJson(req);
      const username = String(body.username || "").trim();
      const password = String(body.password || "");
      const role = body.role === "admin" ? "admin" : "viewer";
      const permissions = normalizePermissions(role, body.permissions);
      if (!username || password.length < 6) {
        sendJson(res, 400, { error: "Username is required and password must be at least 6 characters" });
        return;
      }

      const users = await readUsers();
      if (users.some((user) => user.username.toLowerCase() === username.toLowerCase())) {
        sendJson(res, 409, { error: "Username already exists" });
        return;
      }

      const user = {
        username,
        role,
        permissions,
        password_hash: hashPassword(password),
        created_at: new Date().toISOString(),
      };
      users.push(user);
      await writeUsers(users);
      sendJson(res, 201, { user: publicUser(user) });
      return;
    }
  }

  if (url.pathname.startsWith("/api/users/") && requestUser.role === "admin") {
    const username = decodeURIComponent(url.pathname.slice("/api/users/".length));
    const users = await readUsers();
    const userIndex = users.findIndex((user) => user.username === username);
    if (userIndex < 0) {
      sendJson(res, 404, { error: "User not found" });
      return;
    }

    if (req.method === "PATCH") {
      const body = await readRequestJson(req);
      if (body.role === "admin" || body.role === "viewer") {
        users[userIndex].role = body.role;
      }
      users[userIndex].permissions = normalizePermissions(
        users[userIndex].role,
        Array.isArray(body.permissions) ? body.permissions : users[userIndex].permissions,
      );
      if (typeof body.password === "string" && body.password) {
        if (body.password.length < 6) {
          sendJson(res, 400, { error: "Password must be at least 6 characters" });
          return;
        }
        users[userIndex].password_hash = hashPassword(body.password);
      }
      await writeUsers(users);
      sendJson(res, 200, { user: publicUser(users[userIndex]) });
      return;
    }

    if (req.method === "DELETE") {
      const adminCount = users.filter((user) => user.role === "admin").length;
      if (users[userIndex].role === "admin" && adminCount <= 1) {
        sendJson(res, 400, { error: "Cannot delete the last admin" });
        return;
      }
      users.splice(userIndex, 1);
      await writeUsers(users);
      sendJson(res, 200, { ok: true });
      return;
    }
  }

  if (url.pathname.startsWith("/api/users/") && requestUser.role !== "admin") {
    sendForbidden(res);
    return;
  }

  if (url.pathname === "/api/data") {
    const selectedDate = url.searchParams.get("date");
    const search = String(url.searchParams.get("search") || "").trim().toLowerCase();
    const payload = await readRunData();
    maybeStartRecentPreload(payload);
    const dates = [...new Set(payload.runs.map((run) => run.run_date).filter(Boolean))].sort();
    const activeDate = selectedDate || (dates.includes(localDateIso()) ? localDateIso() : dates.at(-1)) || localDateIso();
    const auto_sync = await maybeStartAutoSync(payload, selectedDate || localDateIso());
    let runs = payload.runs.filter((run) => run.run_date === activeDate);

    if (search) {
      runs = runs.filter((run) => String(run.customer_code || "").toLowerCase().includes(search));
    }

    const [localHydratedRuns, googleHydratedByFolderId] = await Promise.all([
      Promise.all(runs.map(listLocalRunDetails)),
      hydrateGoogleRuns(runs),
    ]);
    const hydratedRuns = localHydratedRuns.map((run) => (
      googleHydratedByFolderId.get(run.folder_id) || run
    ));
    const customerCount = new Set(hydratedRuns.map((run) => run.customer_code)).size;
    const imageCount = hydratedRuns.reduce((total, run) => total + (Array.isArray(run.images) ? run.images.length : 0), 0);

    sendJson(res, 200, {
      dates,
      selected_date: activeDate,
      generated_at: payload.generated_at,
      cache_missing: payload.cache_missing,
      auto_sync,
      parse_errors: payload.parse_errors,
      metrics: {
        customers: customerCount,
        runs: hydratedRuns.length,
        images: imageCount,
      },
      groups: groupByCustomer(hydratedRuns),
    });
    return;
  }

  if (url.pathname === "/api/status") {
    sendJson(res, 200, await readJsonFile(syncStatusPath, {}));
    return;
  }

  if (url.pathname === "/api/rebuild" && req.method === "POST") {
    sendJson(res, startSync("rebuild") ? 202 : 500, { ok: existsSync(syncScriptPath) });
    return;
  }

  if (url.pathname === "/api/refresh-date" && req.method === "POST") {
    const selectedDate = url.searchParams.get("date");
    if (!selectedDate) {
      sendJson(res, 400, { ok: false, error: "date is required" });
      return;
    }
    sendJson(res, startSync("refresh_date", selectedDate) ? 202 : 500, { ok: existsSync(syncScriptPath) });
    return;
  }

  if (url.pathname === "/api/image") {
    const imagePath = url.searchParams.get("id");
    const accountName = url.searchParams.get("account") || "default";
    const forceRefresh = Boolean(url.searchParams.get("retry"));
    if (!imagePath) {
      sendText(res, 400, "Invalid image id");
      return;
    }

    if (!path.isAbsolute(imagePath)) {
      try {
        const imageBytes = await readGoogleImage(imagePath, accountName, { forceRefresh });
        res.writeHead(200, {
          "content-type": url.searchParams.get("mime") || "image/jpeg",
          "cache-control": "private, max-age=300",
        });
        res.end(imageBytes);
      } catch (error) {
        sendText(res, 502, error instanceof Error ? error.message : "Unable to load Google Drive image");
      }
      return;
    }

    const normalized = path.resolve(imagePath);
    if (!existsSync(normalized)) {
      sendText(res, 404, "Image not found");
      return;
    }

    if (!(await isIndexedLocalImage(normalized))) {
      sendText(res, 403, "Image is not part of the shared run index");
      return;
    }

    res.writeHead(200, {
      "content-type": guessMimeType(normalized),
      "cache-control": "private, max-age=300",
    });
    createReadStream(normalized).pipe(res);
    return;
  }

  sendJson(res, 404, { error: "Not found" });
}

async function serveStatic(req, res, url) {
  const requestedPath = decodeURIComponent(url.pathname === "/" ? "/index.html" : url.pathname);
  const resolvedPath = path.resolve(staticRoot, `.${requestedPath}`);
  if (!resolvedPath.startsWith(staticRoot)) {
    sendText(res, 403, "Forbidden");
    return;
  }

  const fallbackPath = path.join(staticRoot, "index.html");
  const filePath = existsSync(resolvedPath) ? resolvedPath : fallbackPath;
  res.writeHead(200, { "content-type": guessMimeType(filePath) });
  createReadStream(filePath).pipe(res);
}

const server = http.createServer((req, res) => {
  const url = new URL(req.url || "/", `http://${req.headers.host || "127.0.0.1"}`);
  if (url.pathname.startsWith("/api/")) {
    handleApi(req, res, url).catch((error) => {
      sendJson(res, 500, { error: error instanceof Error ? error.message : String(error) });
    });
    return;
  }

  serveStatic(req, res, url).catch((error) => {
    sendJson(res, 500, { error: error instanceof Error ? error.message : String(error) });
  });
});

const port = Number(process.env.PORT || 4174);
const host = process.env.HOST || "0.0.0.0";

function lanUrls() {
  const urls = [`http://127.0.0.1:${port}`];
  for (const interfaces of Object.values(os.networkInterfaces())) {
    for (const networkInterface of interfaces || []) {
      if (networkInterface.family !== "IPv4" || networkInterface.internal) {
        continue;
      }
      urls.push(`http://${networkInterface.address}:${port}`);
    }
  }
  return urls;
}

server.on("error", (error) => {
  if (error?.code === "EADDRINUSE") {
    console.log(`SnappySjaak shadow app is already running on port ${port}.`);
    console.log("Try one of these addresses:");
    for (const url of lanUrls()) {
      console.log(`  ${url}`);
    }
    console.log("Close the other shadow app window/process first if you want to restart it.");
    process.exit(0);
  }

  console.error(error);
  process.exit(1);
});

server.listen(port, host, () => {
  console.log("SnappySjaak shadow app is running.");
  console.log("Open on this PC:");
  console.log(`  http://127.0.0.1:${port}`);
  console.log("Open from another PC on the same network:");
  for (const url of lanUrls().filter((url) => !url.includes("127.0.0.1"))) {
    console.log(`  ${url}`);
  }
});
