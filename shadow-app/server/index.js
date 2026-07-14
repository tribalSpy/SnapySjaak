import { createReadStream, existsSync, promises as fs } from "node:fs";
import http from "node:http";
import os from "node:os";
import path from "node:path";
import crypto from "node:crypto";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";
import {
  claimNextLlmJob,
  completeLlmJob,
  createLlmJob,
  dbQuery,
  getFustDatabaseStats,
  getDatabaseStatus,
  getLlmQueueSnapshot,
  failLlmJob,
  initializeDatabase,
  isDatabaseEnabled,
  markFustActionDeletedInDatabase,
  saveFustActionToDatabase,
  upsertLlmAgentHeartbeat,
} from "./db.js";
import { createBunchesService } from "./bunches.js";

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
const ukdocsStatePath = path.join(cacheDir, "ukdocs-state.json");
const ukdocsPrintFilesDir = path.join(cacheDir, "ukdocs-print-files");
const fustBackupDir = path.join(cacheDir, "fust-backups");
const syncScriptPath = path.join(repoRoot, "sync_index.py");
const driveBridgePath = path.join(appRoot, "server", "drive_bridge.py");
const syncWorkerPath = path.join(appRoot, "server", "sync_worker.js");
const halLocationsWorkerPath = path.join(appRoot, "server", "hal_locations_worker.py");
const expeditionStickerWorkerPath = path.join(appRoot, "server", "expedition_sticker_worker.py");
const fustImportWorkerPath = path.join(appRoot, "server", "fust_import_worker.py");
const fustListWorkerPath = path.join(appRoot, "server", "fust_list_worker.py");
const ukdocsWorkerPath = path.join(appRoot, "server", "ukdocs_worker.py");
const ukdocsCsiWorkerPath = path.join(appRoot, "server", "ukdocs_csi_worker.py");
const dagFoutjesHtmlPathCandidates = [
  path.join(repoRoot, "foutjeskoelcel", "bledy-chlodnia (1).html"),
  path.join(process.cwd(), "foutjeskoelcel", "bledy-chlodnia (1).html"),
  path.join(appRoot, "..", "foutjeskoelcel", "bledy-chlodnia (1).html"),
  path.join(appRoot, "foutjeskoelcel", "bledy-chlodnia (1).html"),
  path.join(appRoot, "public", "dag-foutjes.html"),
];
const googleImageCacheDir = path.join(cacheDir, "shadow-google-images");
const googleRunDetailsCacheDir = path.join(cacheDir, "shadow-google-run-details");
const halLocationsCacheDir = path.join(cacheDir, "hal-locations");
const expeditionStickerStatePath = path.join(cacheDir, "expedition-stickers.json");
const expeditionStickerFilesDir = path.join(cacheDir, "expedition-stickers");
const dagFoutjesStatePath = path.join(cacheDir, "dag-foutjes.json");
const bunchesStatePath = path.join(cacheDir, "bunches-state.json");
const bunchesAppHtmlPath = path.join(appRoot, "public", "bunches.html");
const bunchesSeedDir = path.join(appRoot, "server", "bunches-seed");
const usersSeedPathCandidates = [
  process.env.SHADOW_USERS_SEED_PATH,
  process.platform === "win32" ? null : "/etc/secrets/shadow-users.json",
].filter(Boolean);
const staticRoot = existsSync(path.join(appRoot, "dist"))
  ? path.join(appRoot, "dist")
  : path.join(appRoot, "public");
const fustListTemplatePathCandidates = [
  path.join(appRoot, "public", "test fust invoice.xlsx"),
  path.join(staticRoot, "test fust invoice.xlsx"),
];
const autoSyncOnVisit = process.env.AUTO_SYNC_ON_VISIT !== "0";
const autoSyncThrottleMs = Number(process.env.AUTO_SYNC_THROTTLE_MINUTES || 5) * 60 * 1000;
const syncStatusStaleMinutes = Math.max(1, Number(process.env.SHADOW_SYNC_STALE_MINUTES || 30));
const llmPollerApiKey = String(process.env.SHADOW_LLM_POLLER_API_KEY || "").trim();
const autoSyncStartedAt = new Map();
const recentPreloadDays = Math.max(0, Number(process.env.SHADOW_PRELOAD_RECENT_DAYS || 3));
const recentPreloadMaxImagesRaw = Number(process.env.SHADOW_PRELOAD_MAX_IMAGES || 120);
const recentPreloadMaxImages = Number.isFinite(recentPreloadMaxImagesRaw)
  ? recentPreloadMaxImagesRaw
  : 120;
const googleRunDetailsCacheTtlMinutes = Math.max(0, Number(process.env.SHADOW_RUN_DETAILS_TTL_MINUTES || 720));
const recentPreloadStartedAt = new Map();
const sessions = new Map();
const halLocationSessions = new Map();
const sessionCookieName = "snappysjaak_shadow_session";
const halLocationSessionTtlMs = 30 * 60 * 1000;
const allPermissions = [
  "photos:view",
  "fust:view",
  "fust:in",
  "fust:out",
  "fust:overview",
  "fust:manage",
  "cmr:view",
  "hal_locations:view",
  "expedition_stickers:view",
  "bunches:view",
  "fouten_overview:view",
  "cmr:manage",
  "clock:view",
  "clock:manage",
  "users:manage",
  "settings:manage",
  "ukdocs:view",
  "ukdocs_inspection:view",
  "ukdocs_csi:view",
];
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
  FOUTEN_OVERVIEW_VIEW: "fouten_overview:view",
  CMR_MANAGE: "cmr:manage",
  CLOCK_VIEW: "clock:view",
  CLOCK_MANAGE: "clock:manage",
  USERS_MANAGE: "users:manage",
  SETTINGS_MANAGE: "settings:manage",
  UKDOCS_VIEW: "ukdocs:view",
  UKDOCS_INSPECTION_VIEW: "ukdocs_inspection:view",
  UKDOCS_CSI_VIEW: "ukdocs_csi:view",
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
  support_email_recipients: [],
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
  gmail_refresh_token: "",
  gmail_connected_email: "",
  clock_spreadsheet_id: "",
  clock_employee_sheet_name: "badges",
  clock_records_sheet_name: "backup",
  hal_locations_spreadsheet_id: "",
  hal_locations_sheet_name: "ERP_PASTE",
  ukdocs_print_spreadsheet_id: "",
  ukdocs_print_sheet_name: "PD keuringen",
  cmr_default_template_name: "",
  cmr_manage_usernames: [],
};

const defaultUkdocsState = {
  company_settings: {
    company_name: "",
    address: "",
    phone: "",
    email: "",
    website: "",
    vat_number: "",
    eori_number: "",
    chamber_of_commerce_number: "",
    iban: "",
    bic_swift: "",
    rex_registration: "",
    default_footer_text: "",
    preferential_origin_declaration: "",
    logo_name: "",
  },
  export_defaults: {
    destination_country: "GB / United Kingdom",
    regulation: "Export",
    border_transport_mode: "Road",
    border_transport_nationality: "NL",
    customs_office_of_exit: "",
    location: "",
    delivery_terms: "",
    delivery_terms_city: "",
    currency: "GBP",
    freight_costs: "",
    insurance: "",
    importer_field: "",
    vessel_field: "",
    phyto_fields: "",
    kcb_fields: "",
    certificate_fields: "",
    value_tolerance: "0.01",
    weight_tolerance: "0.001",
    quantity_tolerance: "0",
    packages_tolerance: "0",
  },
  templates: {
    invoice_template_name: "",
    export_template_name: "",
    logo_name: "",
  },
  customers: [],
  column_mappings: {
    "508": { aliases: {} },
    "515": { aliases: {} },
    "1000": { aliases: {} },
    "920": { aliases: {} },
  },
  shipments: [],
  audit_reports: [],
  print_collections: [],
};

const UKDOCS_CSI_DOCUMENT_KINDS = new Set(["temp_phyto", "temp_phyto_plants_file", "temp_phyto_plants_xml_file", "ipaffs_file", "ipaffs_plants_file"]);

const defaultExpeditionStickerState = {
  planning_file: null,
  split_file: null,
};

const defaultDagFoutjesState = {
  shared: {},
};

const bunchesService = createBunchesService({
  statePath: bunchesStatePath,
  seedDir: bunchesSeedDir,
});

const cmrPrintDataDirCandidates = [
  path.join(cacheDir, "cmrprint-data"),
  path.join(repoRoot, "cmrprint", "CMRPrint", "Data"),
  path.join(repoRoot, "cmrprint", "CMRPrint", "bin", "Release", "net9.0-windows", "win-x64", "publish", "Data"),
  path.join(repoRoot, "cmrprint", "CMRPrint", "bin", "Release", "net9.0-windows", "win-x64", "Data"),
  path.join(process.cwd(), "cmrprint", "CMRPrint", "Data"),
  path.join(appRoot, "..", "cmrprint", "CMRPrint", "Data"),
  path.join(appRoot, "cmrprint", "CMRPrint", "Data"),
];

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

function readAgentApiKey(req) {
  const directHeader = String(req.headers["x-shadow-agent-key"] || "").trim();
  if (directHeader) {
    return directHeader;
  }
  const authorization = String(req.headers.authorization || "").trim();
  if (authorization.toLowerCase().startsWith("bearer ")) {
    return authorization.slice(7).trim();
  }
  return "";
}

function llmPollerEnabled() {
  return Boolean(llmPollerApiKey);
}

function sendPollerUnauthorized(res) {
  sendJson(res, 401, {
    error: "Invalid or missing poller API key",
    poller_enabled: llmPollerEnabled(),
  });
}

function decodeXmlEntities(value) {
  return String(value || "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#x0D;/g, "\r")
    .replace(/&#10;/g, "\n")
    .replace(/&amp;/g, "&");
}

function extractXmlBlocks(xml, tagName) {
  const matches = [];
  const pattern = new RegExp(`<${tagName}(?:\\s[^>]*)?>([\\s\\S]*?)</${tagName}>`, "g");
  let match = pattern.exec(xml);
  while (match) {
    matches.push(match[1]);
    match = pattern.exec(xml);
  }
  return matches;
}

function extractXmlValue(xml, tagName) {
  const match = new RegExp(`<${tagName}(?:\\s[^>]*)?>([\\s\\S]*?)</${tagName}>`).exec(xml);
  return decodeXmlEntities(match?.[1] || "").trim();
}

function parseCmrPrintFieldAssignments(xml) {
  return extractXmlBlocks(xml, "FieldAssignment").map((block) => ({
    field_name: extractXmlValue(block, "FieldName"),
    value: extractXmlValue(block, "Value"),
  })).filter((item) => item.field_name);
}

function parseCmrPrintProfiles(xml, tagName) {
  return extractXmlBlocks(xml, tagName).map((block) => ({
    name: extractXmlValue(block, "Name"),
    country: extractXmlValue(block, "Country"),
    place: extractXmlValue(block, "Place"),
    field_assignments: parseCmrPrintFieldAssignments(extractXmlValue(block, "FieldAssignments") || block),
  })).filter((item) => item.name || item.field_assignments.length);
}

function parseCmrPrintCustomers(xml) {
  return extractXmlBlocks(xml, "Customer").map((block) => ({
    name: extractXmlValue(block, "Name"),
    address: extractXmlValue(block, "Address"),
    city: extractXmlValue(block, "City"),
    country: extractXmlValue(block, "Country"),
    vat_number: extractXmlValue(block, "VatNumber"),
    exporter_profile_name: extractXmlValue(block, "ExporterProfileName"),
    transport_profile_name: extractXmlValue(block, "TransportProfileName"),
    loading_place_profile_name: extractXmlValue(block, "LoadingPlaceProfileName"),
    place_of_issue: extractXmlValue(block, "PlaceOfIssue"),
    field_assignments: parseCmrPrintFieldAssignments(extractXmlValue(block, "FieldAssignments") || block),
  })).filter((item) => item.name || item.address || item.field_assignments.length);
}

function parseCmrPrintTemplateIntEntries(xml, tagName) {
  return extractXmlBlocks(xml, tagName).map((block) => ({
    field_name: extractXmlValue(block, "FieldName"),
    value: Number(extractXmlValue(block, "Value") || 0),
  })).filter((item) => item.field_name);
}

function parseCmrPrintTemplatePointEntries(xml, tagName) {
  return extractXmlBlocks(xml, tagName).map((block) => ({
    field_name: extractXmlValue(block, "FieldName"),
    x: Number(extractXmlValue(block, "X") || 0),
    y: Number(extractXmlValue(block, "Y") || 0),
  })).filter((item) => item.field_name);
}

function cmrPrintPlaces() {
  return [
    { place_number: 1, field_name: "ConsignorName", description: "1. Sender", default_x: 40, default_y: 80, default_font_size: 9 },
    { place_number: 2, field_name: "ConsignorDetails", description: "2. Destination", default_x: 40, default_y: 130, default_font_size: 8 },
    { place_number: 3, field_name: "LoadingInstructions", description: "3. Place of Delivery Good", default_x: 40, default_y: 160, default_font_size: 8 },
    { place_number: 4, field_name: "ConsignorRemarks", description: "4. Place and Date of Reception", default_x: 40, default_y: 200, default_font_size: 8 },
    { place_number: 5, field_name: "DocumentsAttached", description: "5. Documents attached", default_x: 40, default_y: 240, default_font_size: 8 },
    { place_number: 6, field_name: "Seals", description: "6. Marks and Numbers", default_x: 120, default_y: 240, default_font_size: 8 },
    { place_number: 7, field_name: "PackagingType", description: "7. Number of Packages", default_x: 200, default_y: 240, default_font_size: 8 },
    { place_number: 8, field_name: "GoodsDescription", description: "8. Goods description", default_x: 40, default_y: 280, default_font_size: 9 },
    { place_number: 9, field_name: "NatureofGoods", description: "9. Nature of Goods", default_x: 40, default_y: 360, default_font_size: 9 },
    { place_number: 10, field_name: "LoadingOrderNumber", description: "10. Statistical Number", default_x: 200, default_y: 360, default_font_size: 8 },
    { place_number: 11, field_name: "TransportChargesPlace", description: "11. Gross Weight", default_x: 400, default_y: 80, default_font_size: 8 },
    { place_number: 12, field_name: "ConsigeeName", description: "12. Volume in m3", default_x: 400, default_y: 110, default_font_size: 9 },
    { place_number: 13, field_name: "ConsigneeDetails", description: "13. Sender Instructions", default_x: 400, default_y: 160, default_font_size: 8 },
    { place_number: 14, field_name: "UnloadingInstructions", description: "14. Instructions regarding Payment", default_x: 400, default_y: 190, default_font_size: 8 },
    { place_number: 15, field_name: "CarrierRemarks", description: "15. Cash on Delivery", default_x: 40, default_y: 430, default_font_size: 8 },
    { place_number: 16, field_name: "ConsigneeRemarks", description: "16. Carrier", default_x: 400, default_y: 430, default_font_size: 8 },
    { place_number: 17, field_name: "TransportAuthorizations", description: "17. Successive Carriers", default_x: 40, default_y: 500, default_font_size: 8 },
    { place_number: 18, field_name: "RouteInfo", description: "18. Carrier Observations", default_x: 200, default_y: 500, default_font_size: 8 },
    { place_number: 19, field_name: "InsuranceRemarks", description: "19. Special Agreements", default_x: 40, default_y: 540, default_font_size: 8 },
    { place_number: 20, field_name: "CarrierSignature", description: "20. To be Paid By", default_x: 40, default_y: 580, default_font_size: 8 },
    { place_number: 21, field_name: "ExportDate", description: "21. Export/transport date", default_x: 400, default_y: 540, default_font_size: 9 },
    { place_number: 22, field_name: "SignaturePlace1", description: "22. Signature Sender", default_x: 40, default_y: 620, default_font_size: 8 },
    { place_number: 23, field_name: "SignaturePlace2", description: "23. Signature of the carrier", default_x: 270, default_y: 620, default_font_size: 8 },
    { place_number: 24, field_name: "SignaturePlace3", description: "24. Signature Good received", default_x: 500, default_y: 620, default_font_size: 8 },
  ];
}

function parseCmrPrintTemplate(xml, filename) {
  return {
    name: extractXmlValue(xml, "Name") || filename.replace(/\.xml$/i, ""),
    created_date: extractXmlValue(xml, "CreatedDate"),
    font_sizes: parseCmrPrintTemplateIntEntries(extractXmlValue(xml, "FontSizeEntries") || xml, "TemplateIntSetting"),
    vertical_offsets: parseCmrPrintTemplateIntEntries(extractXmlValue(xml, "VerticalOffsetEntries") || xml, "TemplateIntSetting"),
    field_positions: parseCmrPrintTemplatePointEntries(extractXmlValue(xml, "FieldPositionEntries") || xml, "TemplatePointSetting"),
    field_widths: parseCmrPrintTemplateIntEntries(extractXmlValue(xml, "FieldWidthEntries") || xml, "TemplateIntSetting"),
    field_heights: parseCmrPrintTemplateIntEntries(extractXmlValue(xml, "FieldHeightEntries") || xml, "TemplateIntSetting"),
    source_file: filename,
  };
}

function cmrPrintCandidateStatus() {
  return [...new Set(cmrPrintDataDirCandidates.map((candidate) => path.resolve(candidate)))].map((candidate) => ({
    path: candidate,
    exists: existsSync(candidate),
    has_app_data: existsSync(path.join(candidate, "app-data.xml")),
    has_templates_dir: existsSync(path.join(candidate, "Templates")),
  }));
}

function resolveCmrPrintDataDir() {
  const candidates = cmrPrintCandidateStatus();
  const exactMatch = candidates.find((candidate) => candidate.has_app_data || candidate.has_templates_dir);
  return {
    dataDir: exactMatch?.path || "",
    candidates,
  };
}

async function loadCmrPrintData() {
  const candidates = cmrPrintCandidateStatus();
  const { dataDir } = await ensureCmrPrintPersistentDataDir();
  if (!dataDir) {
    return {
      available: false,
      data_dir: "",
      templates_dir: "",
      customers: [],
      exporters: [],
      transport_infos: [],
      loading_places: [],
      templates: [],
      places: cmrPrintPlaces(),
      debug_candidates: candidates,
    };
  }

  const appDataPath = path.join(dataDir, "app-data.xml");
  const templatesDir = path.join(dataDir, "Templates");
  const appDataXml = existsSync(appDataPath) ? await fs.readFile(appDataPath, "utf8") : "";
  const customers = parseCmrPrintCustomers(extractXmlValue(appDataXml, "Customers") || appDataXml).sort((left, right) => left.name.localeCompare(right.name));
  const exporters = parseCmrPrintProfiles(extractXmlValue(appDataXml, "Exporters") || appDataXml, "ProfileRecord").sort((left, right) => left.name.localeCompare(right.name));
  const transportInfos = parseCmrPrintProfiles(extractXmlValue(appDataXml, "TransportInfos") || appDataXml, "ProfileRecord").sort((left, right) => left.name.localeCompare(right.name));
  const loadingPlaces = parseCmrPrintProfiles(extractXmlValue(appDataXml, "LoadingPlaces") || appDataXml, "ProfileRecord").sort((left, right) => left.name.localeCompare(right.name));
  const templates = [];
  if (existsSync(templatesDir)) {
    const entries = await fs.readdir(templatesDir, { withFileTypes: true });
    for (const entry of entries) {
      if (!entry.isFile() || path.extname(entry.name).toLowerCase() !== ".xml") {
        continue;
      }
      const xml = await fs.readFile(path.join(templatesDir, entry.name), "utf8");
      templates.push(parseCmrPrintTemplate(xml, entry.name));
    }
  }
  templates.sort((left, right) => left.name.localeCompare(right.name));
  return {
    available: true,
    data_dir: dataDir,
    templates_dir: templatesDir,
    customers,
    exporters,
    transport_infos: transportInfos,
    loading_places: loadingPlaces,
    templates,
    places: cmrPrintPlaces(),
    debug_candidates: candidates,
  };
}

function cmrPrintPrimaryDataDir() {
  return path.join(cacheDir, "cmrprint-data");
}

async function copyCmrPrintDirectoryIfMissing(sourceDir, targetDir) {
  if (!sourceDir || !targetDir || path.resolve(sourceDir) === path.resolve(targetDir)) {
    return;
  }
  await fs.mkdir(targetDir, { recursive: true });
  const sourceAppDataPath = path.join(sourceDir, "app-data.xml");
  const targetAppDataPath = path.join(targetDir, "app-data.xml");
  if (existsSync(sourceAppDataPath) && !existsSync(targetAppDataPath)) {
    await fs.copyFile(sourceAppDataPath, targetAppDataPath);
  }
  const sourceTemplatesDir = path.join(sourceDir, "Templates");
  const targetTemplatesDir = path.join(targetDir, "Templates");
  if (existsSync(sourceTemplatesDir)) {
    await fs.mkdir(targetTemplatesDir, { recursive: true });
    const entries = await fs.readdir(sourceTemplatesDir, { withFileTypes: true });
    for (const entry of entries) {
      if (!entry.isFile()) {
        continue;
      }
      const sourcePath = path.join(sourceTemplatesDir, entry.name);
      const targetPath = path.join(targetTemplatesDir, entry.name);
      if (!existsSync(targetPath)) {
        await fs.copyFile(sourcePath, targetPath);
      }
    }
  }
}

async function countCmrTemplateFiles(dataDir) {
  const templatesDir = path.join(dataDir, "Templates");
  if (!existsSync(templatesDir)) {
    return 0;
  }
  const entries = await fs.readdir(templatesDir, { withFileTypes: true });
  return entries.filter((entry) => entry.isFile() && path.extname(entry.name).toLowerCase() === ".xml").length;
}

async function ensureCmrPrintPersistentDataDir() {
  const persistentDataDir = cmrPrintPrimaryDataDir();
  const persistentTemplatesDir = path.join(persistentDataDir, "Templates");
  await fs.mkdir(persistentTemplatesDir, { recursive: true });
  const persistentHasAppData = existsSync(path.join(persistentDataDir, "app-data.xml"));
  const persistentTemplateCount = await countCmrTemplateFiles(persistentDataDir);
  const bootstrapSource = cmrPrintCandidateStatus()
    .map((candidate) => candidate.path)
    .find((candidate) => path.resolve(candidate) !== path.resolve(persistentDataDir)
      && (existsSync(path.join(candidate, "app-data.xml")) || existsSync(path.join(candidate, "Templates"))));
  if (bootstrapSource && (!persistentHasAppData || persistentTemplateCount === 0)) {
    await copyCmrPrintDirectoryIfMissing(bootstrapSource, persistentDataDir);
  }
  return {
    dataDir: persistentDataDir,
    templatesDir: persistentTemplatesDir,
    appDataPath: path.join(persistentDataDir, "app-data.xml"),
  };
}

async function ensureCmrPrintDataDir() {
  return ensureCmrPrintPersistentDataDir();
}

function xmlEscape(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;")
    .replace(/\r/g, "&#x0D;");
}

function serializeFieldAssignments(assignments) {
  const rows = Array.isArray(assignments) ? assignments : [];
  return rows.map((item) => `
      <FieldAssignment>
        <FieldName>${xmlEscape(item?.field_name)}</FieldName>
        <Value>${xmlEscape(item?.value)}</Value>
      </FieldAssignment>`).join("");
}

function serializeProfileRecord(tagName, item) {
  return `
    <${tagName}>
      <Name>${xmlEscape(item?.name)}</Name>
      <Country>${xmlEscape(item?.country)}</Country>
      <Place>${xmlEscape(item?.place)}</Place>
      <FieldAssignments>${serializeFieldAssignments(item?.field_assignments)}
      </FieldAssignments>
    </${tagName}>`;
}

function serializeCustomer(item) {
  return `
    <Customer>
      <Name>${xmlEscape(item?.name)}</Name>
      <Address>${xmlEscape(item?.address)}</Address>
      <City>${xmlEscape(item?.city)}</City>
      <Country>${xmlEscape(item?.country)}</Country>
      <VatNumber>${xmlEscape(item?.vat_number)}</VatNumber>
      <ExporterProfileName>${xmlEscape(item?.exporter_profile_name)}</ExporterProfileName>
      <TransportProfileName>${xmlEscape(item?.transport_profile_name)}</TransportProfileName>
      <LoadingPlaceProfileName>${xmlEscape(item?.loading_place_profile_name)}</LoadingPlaceProfileName>
      <PlaceOfIssue>${xmlEscape(item?.place_of_issue)}</PlaceOfIssue>
      <FieldAssignments>${serializeFieldAssignments(item?.field_assignments)}
      </FieldAssignments>
    </Customer>`;
}

function normalizeCmrAssignments(assignments) {
  if (!Array.isArray(assignments)) {
    return [];
  }
  return assignments.map((item) => ({
    field_name: String(item?.field_name || "").trim(),
    value: String(item?.value || ""),
  })).filter((item) => item.field_name);
}

function normalizeCmrProfile(item) {
  return {
    name: String(item?.name || "").trim(),
    country: String(item?.country || "").trim(),
    place: String(item?.place || "").trim(),
    field_assignments: normalizeCmrAssignments(item?.field_assignments),
  };
}

function normalizeCmrCustomer(item) {
  return {
    name: String(item?.name || "").trim(),
    address: String(item?.address || ""),
    city: String(item?.city || ""),
    country: String(item?.country || "").trim(),
    vat_number: String(item?.vat_number || ""),
    exporter_profile_name: String(item?.exporter_profile_name || ""),
    transport_profile_name: String(item?.transport_profile_name || ""),
    loading_place_profile_name: String(item?.loading_place_profile_name || ""),
    place_of_issue: String(item?.place_of_issue || ""),
    field_assignments: normalizeCmrAssignments(item?.field_assignments),
  };
}

function exporterBlockFromCmrProfile(profile) {
  const fieldAssignments = Array.isArray(profile?.field_assignments) ? profile.field_assignments : [];
  const consignorValue = fieldAssignments.find((item) => item?.field_name === "ConsignorName")?.value || "";
  const signatureValue = fieldAssignments.find((item) => item?.field_name === "SignaturePlace1")?.value || "";
  const block = String(consignorValue || signatureValue || "").trim();
  return {
    name: String(profile?.name || "").trim(),
    country: String(profile?.country || "").trim(),
    place: String(profile?.place || "").trim(),
    block,
  };
}

async function saveCmrPrintAppData(payload) {
  const { appDataPath } = await ensureCmrPrintDataDir();
  const customers = (Array.isArray(payload?.customers) ? payload.customers : []).map(normalizeCmrCustomer).filter((item) => item.name);
  const exporters = (Array.isArray(payload?.exporters) ? payload.exporters : []).map(normalizeCmrProfile).filter((item) => item.name);
  const transportInfos = (Array.isArray(payload?.transport_infos) ? payload.transport_infos : []).map(normalizeCmrProfile).filter((item) => item.name);
  const loadingPlaces = (Array.isArray(payload?.loading_places) ? payload.loading_places : []).map(normalizeCmrProfile).filter((item) => item.name);
  const xml = `<?xml version="1.0" encoding="utf-8"?>
<AppDataStore>
  <Customers>${customers.map(serializeCustomer).join("")}
  </Customers>
  <Exporters>${exporters.map((item) => serializeProfileRecord("ProfileRecord", item)).join("")}
  </Exporters>
  <TransportInfos>${transportInfos.map((item) => serializeProfileRecord("ProfileRecord", item)).join("")}
  </TransportInfos>
  <LoadingPlaces>${loadingPlaces.map((item) => serializeProfileRecord("ProfileRecord", item)).join("")}
  </LoadingPlaces>
</AppDataStore>
`;
  await fs.writeFile(appDataPath, xml, "utf8");
}

function normalizeCmrTemplatePayload(template) {
  return {
    name: String(template?.name || "").trim(),
    created_date: String(template?.created_date || new Date().toISOString()),
    font_sizes: Array.isArray(template?.font_sizes) ? template.font_sizes : [],
    vertical_offsets: Array.isArray(template?.vertical_offsets) ? template.vertical_offsets : [],
    field_positions: Array.isArray(template?.field_positions) ? template.field_positions : [],
    field_widths: Array.isArray(template?.field_widths) ? template.field_widths : [],
    field_heights: Array.isArray(template?.field_heights) ? template.field_heights : [],
  };
}

function serializeTemplateIntEntries(entries) {
  return entries.map((entry) => `
    <TemplateIntSetting>
      <FieldName>${xmlEscape(entry?.field_name)}</FieldName>
      <Value>${Number(entry?.value || 0)}</Value>
    </TemplateIntSetting>`).join("");
}

function serializeTemplatePointEntries(entries) {
  return entries.map((entry) => `
    <TemplatePointSetting>
      <FieldName>${xmlEscape(entry?.field_name)}</FieldName>
      <X>${Number(entry?.x || 0)}</X>
      <Y>${Number(entry?.y || 0)}</Y>
    </TemplatePointSetting>`).join("");
}

async function saveCmrPrintTemplate(template) {
  const normalized = normalizeCmrTemplatePayload(template);
  if (!normalized.name) {
    throw new Error("Template name is required");
  }
  const { templatesDir } = await ensureCmrPrintDataDir();
  const filePath = path.join(templatesDir, `${normalized.name}.xml`);
  const xml = `<?xml version="1.0" encoding="utf-8"?>
<CmrTemplate>
  <Name>${xmlEscape(normalized.name)}</Name>
  <CreatedDate>${xmlEscape(normalized.created_date)}</CreatedDate>
  <FontSizeEntries>${serializeTemplateIntEntries(normalized.font_sizes)}
  </FontSizeEntries>
  <VerticalOffsetEntries>${serializeTemplateIntEntries(normalized.vertical_offsets)}
  </VerticalOffsetEntries>
  <FieldPositionEntries>${serializeTemplatePointEntries(normalized.field_positions)}
  </FieldPositionEntries>
  <FieldWidthEntries>${serializeTemplateIntEntries(normalized.field_widths)}
  </FieldWidthEntries>
  <FieldHeightEntries>${serializeTemplateIntEntries(normalized.field_heights)}
  </FieldHeightEntries>
</CmrTemplate>
`;
  await fs.writeFile(filePath, xml, "utf8");
  return filePath;
}

async function deleteCmrPrintTemplate(templateName) {
  const safeName = String(templateName || "").trim();
  if (!safeName) {
    throw new Error("Template name is required");
  }
  const { templatesDir } = await ensureCmrPrintDataDir();
  const filePath = path.join(templatesDir, `${safeName}.xml`);
  if (existsSync(filePath)) {
    await fs.unlink(filePath);
  }
}

function canManageCmrWorkspace(user, settings) {
  if (!user) {
    return false;
  }
  const permissions = normalizePermissions(user.role, user.permissions);
  if (user.role === "admin" || permissions.includes(PERMISSIONS.SETTINGS_MANAGE) || permissions.includes(PERMISSIONS.CMR_MANAGE)) {
    return true;
  }
  return settings.cmr_manage_usernames.includes(String(user.username || "").trim().toLowerCase());
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
    ".pdf": "application/pdf",
    ".xls": "application/vnd.ms-excel",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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
  const values = Array.isArray(recipients)
    ? recipients
    : String(recipients || "")
      .split(/[\n,;]+/);

  return [...new Set(
    values
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

function normalizeFustConfirmationReminder(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return {
      last_sent_at: "",
      last_sent_for_date: "",
      last_attempt_at: "",
      last_attempt_for_date: "",
      sent_count: 0,
      last_error: "",
    };
  }
  return {
    last_sent_at: String(value?.last_sent_at || "").trim(),
    last_sent_for_date: String(value?.last_sent_for_date || "").trim(),
    last_attempt_at: String(value?.last_attempt_at || "").trim(),
    last_attempt_for_date: String(value?.last_attempt_for_date || "").trim(),
    sent_count: Number(value?.sent_count || 0) || 0,
    last_error: String(value?.last_error || "").trim(),
  };
}

function normalizeFustImportSource(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return null;
  }
  const importKey = String(value?.import_key || "").trim();
  const fileName = String(value?.file_name || "").trim();
  const sheetName = String(value?.sheet_name || "").trim();
  const rowNumber = Number(value?.row_number || 0);
  const sourceDate = String(value?.source_date || "").trim();
  const importedAt = String(value?.imported_at || "").trim();
  const importedBy = String(value?.imported_by || "").trim();
  const importKind = String(value?.import_kind || "").trim();
  if (!importKey && !fileName && !sheetName && !rowNumber && !sourceDate && !importedAt && !importedBy && !importKind) {
    return null;
  }
  return {
    import_key: importKey,
    file_name: fileName,
    sheet_name: sheetName,
    row_number: rowNumber,
    source_date: sourceDate,
    imported_at: importedAt,
    imported_by: importedBy,
    import_kind: importKind,
  };
}

function normalizeCmrManageUsernames(values) {
  if (!Array.isArray(values)) {
    return [];
  }
  return [...new Set(values.map((value) => String(value || "").trim().toLowerCase()).filter(Boolean))];
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
    support_email_recipients: normalizeEmailRecipients(settings?.support_email_recipients),
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
    gmail_refresh_token: String(settings?.gmail_refresh_token || ""),
    gmail_connected_email: String(settings?.gmail_connected_email || "").trim(),
    clock_spreadsheet_id: String(settings?.clock_spreadsheet_id || "").trim(),
    clock_employee_sheet_name: String(settings?.clock_employee_sheet_name || defaultFustSettings.clock_employee_sheet_name).trim() || defaultFustSettings.clock_employee_sheet_name,
    clock_records_sheet_name: String(settings?.clock_records_sheet_name || defaultFustSettings.clock_records_sheet_name).trim() || defaultFustSettings.clock_records_sheet_name,
    hal_locations_spreadsheet_id: String(settings?.hal_locations_spreadsheet_id || settings?.spreadsheet_id || "").trim(),
    hal_locations_sheet_name: String(settings?.hal_locations_sheet_name || defaultFustSettings.hal_locations_sheet_name).trim() || defaultFustSettings.hal_locations_sheet_name,
    ukdocs_print_spreadsheet_id: String(settings?.ukdocs_print_spreadsheet_id || "").trim(),
    ukdocs_print_sheet_name: String(settings?.ukdocs_print_sheet_name || defaultFustSettings.ukdocs_print_sheet_name).trim() || defaultFustSettings.ukdocs_print_sheet_name,
    cmr_default_template_name: String(settings?.cmr_default_template_name || "").trim(),
    cmr_manage_usernames: normalizeCmrManageUsernames(settings?.cmr_manage_usernames),
  };
}

async function readFustSettings() {
  const payload = await readJsonFile(fustSettingsPath, defaultFustSettings);
  return normalizeFustSettings(payload);
}

async function writeFustSettings(settings) {
  await writeJsonFile(fustSettingsPath, normalizeFustSettings(settings));
}

async function sendSupportAttentionEmail(settings, details = {}) {
  const recipients = normalizeEmailRecipients(settings?.support_email_recipients);
  if (!recipients.length) {
    return { ok: false, reason: "No ICT support recipients configured" };
  }
  if (!settings?.smtp_host || !settings?.smtp_username || !settings?.smtp_password || !settings?.smtp_from) {
    return { ok: false, reason: "SMTP is not fully configured" };
  }

  const subject = [
    "Support needed",
    String(details.service || "Connection"),
    String(details.action || "").trim(),
  ].filter(Boolean).join(" | ");

  const body = [
    "A SnappySjaak service needs attention.",
    "",
    `Service: ${String(details.service || "-")}`,
    `Action attempted: ${String(details.action || "-")}`,
    `User: ${String(details.username || "-")}`,
    `Time: ${new Date().toISOString()}`,
    `Connected account / target: ${String(details.connected_account || "-")}`,
    `What failed: ${String(details.error || "-")}`,
    "",
    `What must be connected: ${String(details.reconnect_target || "-")}`,
    `Workaround: ${String(details.workaround || "Open a private browser window, sign in with the required account, and reconnect the service.")}`,
    "",
    `Request path: ${String(details.path || "-")}`,
  ].join("\n");

  try {
    await runPythonBridge(
      ["email-send"],
      JSON.stringify({
        recipients,
        subject,
        body,
        attachments: [],
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
    return { ok: true };
  } catch (error) {
    return { ok: false, reason: error instanceof Error ? error.message : String(error) };
  }
}

function buildReconnectGuidance(serviceLabel, connectedAccount, fallbackTarget) {
  const exactTarget = String(connectedAccount || fallbackTarget || "").trim();
  if (exactTarget) {
    return `Reconnect ${serviceLabel} with ${exactTarget}. If that Google account is not visible in this browser, open a private / incognito window and sign in with ${exactTarget}.`;
  }
  return `Reconnect ${serviceLabel}. If the correct Google account is not visible in this browser, open a private / incognito window and sign in with the required account first.`;
}

function normalizeUkdocsText(value) {
  return String(value || "").trim();
}

function isHonselersdijkStockControl(collection) {
  const city = normalizeUkdocsText(collection?.city_name).toUpperCase();
  return city === "HONSELERSDIJK";
}

function normalizeUkdocsAliases(aliases) {
  if (!aliases || typeof aliases !== "object") {
    return {};
  }
  return Object.fromEntries(
    Object.entries(aliases)
      .map(([key, value]) => [normalizeUkdocsText(key), Array.isArray(value)
        ? value.map((item) => normalizeUkdocsText(item)).filter(Boolean)
        : String(value || "")
          .split(/\r?\n|,|;/)
          .map((item) => normalizeUkdocsText(item))
          .filter(Boolean)])
      .filter(([key]) => key),
  );
}

function normalizeUkdocsCompanySettings(settings) {
  return {
    company_name: normalizeUkdocsText(settings?.company_name),
    address: String(settings?.address || "").trim(),
    phone: normalizeUkdocsText(settings?.phone),
    email: normalizeUkdocsText(settings?.email),
    website: normalizeUkdocsText(settings?.website),
    vat_number: normalizeUkdocsText(settings?.vat_number),
    eori_number: normalizeUkdocsText(settings?.eori_number),
    chamber_of_commerce_number: normalizeUkdocsText(settings?.chamber_of_commerce_number),
    iban: normalizeUkdocsText(settings?.iban),
    bic_swift: normalizeUkdocsText(settings?.bic_swift),
    rex_registration: normalizeUkdocsText(settings?.rex_registration),
    default_footer_text: String(settings?.default_footer_text || "").trim(),
    preferential_origin_declaration: String(settings?.preferential_origin_declaration || "").trim(),
    logo_name: normalizeUkdocsText(settings?.logo_name),
  };
}

function normalizeUkdocsCustomer(customer) {
  return {
    id: normalizeUkdocsText(customer?.id) || crypto.randomUUID(),
    customer_name: normalizeUkdocsText(customer?.customer_name),
    match_hub_code: normalizeUkdocsText(customer?.match_hub_code),
    match_remark: String(customer?.match_remark || "").trim(),
    required_phyto: customer?.required_phyto !== false,
    required_export_extra: customer?.required_export_extra === true,
    required_generated_export: customer?.required_generated_export !== false,
    required_generated_invoices: customer?.required_generated_invoices !== false,
    reinspection_uses_email_sync: customer?.reinspection_uses_email_sync === true,
    menu_show_ukdocscsi: customer?.menu_show_ukdocscsi !== false,
    menu_show_ukdocsinspection_inspection_list: customer?.menu_show_ukdocsinspection_inspection_list !== false,
    menu_show_ukdocsinspection_locations_file: customer?.menu_show_ukdocsinspection_locations_file !== false,
    menu_show_ukdocsinspection_phyto: customer?.menu_show_ukdocsinspection_phyto === true,
    menu_show_ukdocsinspection_export_extra: customer?.menu_show_ukdocsinspection_export_extra === true,
    menu_show_ukdocsinspection_generated_invoices: customer?.menu_show_ukdocsinspection_generated_invoices === true,
    menu_show_ukdocsinspection_generated_export: customer?.menu_show_ukdocsinspection_generated_export === true,
    menu_show_ukdocsprint_phyto: customer?.menu_show_ukdocsprint_phyto !== false,
    menu_show_ukdocsprint_export_extra: customer?.menu_show_ukdocsprint_export_extra !== false,
    menu_show_ukdocsprint_generated_invoices: customer?.menu_show_ukdocsprint_generated_invoices !== false,
    menu_show_ukdocsprint_generated_export: customer?.menu_show_ukdocsprint_generated_export === true,
    menu_show_ukdocsprint_inspection_list: customer?.menu_show_ukdocsprint_inspection_list === true,
    menu_show_ukdocsprint_locations_file: customer?.menu_show_ukdocsprint_locations_file === true,
    customer_address: String(customer?.customer_address || "").trim(),
    vat_number: normalizeUkdocsText(customer?.vat_number),
    eori_number: normalizeUkdocsText(customer?.eori_number),
    importer_number: normalizeUkdocsText(customer?.importer_number),
    default_owner: normalizeUkdocsText(customer?.default_owner),
    default_importer: normalizeUkdocsText(customer?.default_importer),
    default_delivery_terms: normalizeUkdocsText(customer?.default_delivery_terms),
    default_city: normalizeUkdocsText(customer?.default_city),
    default_uk_arrival_port: normalizeUkdocsText(customer?.default_uk_arrival_port),
    default_currency: normalizeUkdocsText(customer?.default_currency),
    ready_email_subject: String(customer?.ready_email_subject || "").trim(),
    ready_email_body: String(customer?.ready_email_body || "").trim(),
    csi_email_recipients: normalizeEmailRecipients(customer?.csi_email_recipients),
    csi_email_subject: String(customer?.csi_email_subject || "").trim(),
    csi_email_body: String(customer?.csi_email_body || "").trim(),
    default_invoice_language_text: String(customer?.default_invoice_language_text || "").trim(),
    default_document_references: String(customer?.default_document_references || "").trim(),
    show_invoice_vat_number: customer?.show_invoice_vat_number !== false,
    show_invoice_eori_number: customer?.show_invoice_eori_number !== false,
    show_invoice_importer_number: customer?.show_invoice_importer_number !== false,
    export_defaults: normalizeUkdocsExportDefaults(customer?.export_defaults || {}),
  };
}

function normalizeUkdocsCsiReport(report) {
  if (!report || typeof report !== "object") {
    return {
      status: "",
      job_id: "",
      queued_at: "",
      started_at: "",
      completed_at: "",
      error: "",
      summary: "",
      overall_status: "",
      checks: [],
      products: [],
      flower_products: [],
      plant_products: [],
      source_rows: [],
      manual_checks: [],
      notes: [],
      llm_content: "",
      llm_parse_source: "",
      llm_parse_error: "",
      llm_raw_result_json: "",
    };
  }
  return {
    status: normalizeUkdocsText(report.status),
    job_id: normalizeUkdocsText(report.job_id),
    queued_at: normalizeUkdocsText(report.queued_at),
    started_at: normalizeUkdocsText(report.started_at),
    completed_at: normalizeUkdocsText(report.completed_at),
    error: String(report.error || "").trim(),
    summary: String(report.summary || "").trim(),
    overall_status: normalizeUkdocsText(report.overall_status),
    checks: Array.isArray(report.checks) ? report.checks : [],
    products: Array.isArray(report.products) ? report.products : [],
    flower_products: Array.isArray(report.flower_products) ? report.flower_products : [],
    plant_products: Array.isArray(report.plant_products) ? report.plant_products : [],
    source_rows: Array.isArray(report.source_rows) ? report.source_rows : [],
    manual_checks: Array.isArray(report.manual_checks) ? report.manual_checks.map((item) => String(item || "").trim()).filter(Boolean) : [],
    notes: Array.isArray(report.notes) ? report.notes.map((item) => String(item || "").trim()).filter(Boolean) : [],
    llm_content: String(report.llm_content || "").trim(),
    llm_parse_source: String(report.llm_parse_source || "").trim(),
    llm_parse_error: String(report.llm_parse_error || "").trim(),
    llm_raw_result_json: String(report.llm_raw_result_json || "").trim(),
  };
}

function createEmptyUkdocsCsiReport() {
  return normalizeUkdocsCsiReport(null);
}

function shouldResetUkdocsCsiReportForDocumentKind(kind) {
  return [
    "generated",
    "temp_phyto",
    "temp_phyto_plants_file",
    "temp_phyto_plants_xml_file",
    "ipaffs_file",
    "ipaffs_plants_file",
  ].includes(String(kind || "").trim());
}

function normalizeUkdocsExportDefaults(settings) {
  return {
    destination_country: normalizeUkdocsText(settings?.destination_country) || defaultUkdocsState.export_defaults.destination_country,
    regulation: normalizeUkdocsText(settings?.regulation) || defaultUkdocsState.export_defaults.regulation,
    border_transport_mode: normalizeUkdocsText(settings?.border_transport_mode) || defaultUkdocsState.export_defaults.border_transport_mode,
    border_transport_nationality: normalizeUkdocsText(settings?.border_transport_nationality) || defaultUkdocsState.export_defaults.border_transport_nationality,
    customs_office_of_exit: normalizeUkdocsText(settings?.customs_office_of_exit),
    location: normalizeUkdocsText(settings?.location),
    delivery_terms: normalizeUkdocsText(settings?.delivery_terms),
    delivery_terms_city: normalizeUkdocsText(settings?.delivery_terms_city),
    currency: normalizeUkdocsText(settings?.currency) || defaultUkdocsState.export_defaults.currency,
    freight_costs: normalizeUkdocsText(settings?.freight_costs),
    insurance: normalizeUkdocsText(settings?.insurance),
    importer_field: String(settings?.importer_field || "").trim(),
    vessel_field: String(settings?.vessel_field || "").trim(),
    phyto_fields: String(settings?.phyto_fields || "").trim(),
    kcb_fields: String(settings?.kcb_fields || "").trim(),
    certificate_fields: String(settings?.certificate_fields || "").trim(),
    value_tolerance: normalizeUkdocsText(settings?.value_tolerance) || defaultUkdocsState.export_defaults.value_tolerance,
    weight_tolerance: normalizeUkdocsText(settings?.weight_tolerance) || defaultUkdocsState.export_defaults.weight_tolerance,
    quantity_tolerance: normalizeUkdocsText(settings?.quantity_tolerance) || defaultUkdocsState.export_defaults.quantity_tolerance,
    packages_tolerance: normalizeUkdocsText(settings?.packages_tolerance) || defaultUkdocsState.export_defaults.packages_tolerance,
  };
}

function normalizeUkdocsTemplates(templates) {
  return {
    invoice_template_name: normalizeUkdocsText(templates?.invoice_template_name),
    export_template_name: normalizeUkdocsText(templates?.export_template_name),
    logo_name: normalizeUkdocsText(templates?.logo_name),
  };
}

function normalizeUkdocsColumnMappings(mappings) {
  const source = mappings && typeof mappings === "object" ? mappings : {};
  const next = {};
  for (const category of ["508", "515", "1000", "920"]) {
    next[category] = { aliases: normalizeUkdocsAliases(source?.[category]?.aliases) };
  }
  return next;
}

function normalizeUkdocsInvoiceNumbersByCategory(values) {
  const source = values && typeof values === "object" ? values : {};
  const next = {};
  for (const category of ["508", "515", "1000", "920"]) {
    next[category] = normalizeUkdocsText(source?.[category]);
  }
  return next;
}

function normalizeUkdocsUploadedFiles(files) {
  const source = files && typeof files === "object" ? files : {};
  const next = {};
  for (const category of ["508", "515", "1000", "920"]) {
    const file = source?.[category] || {};
    next[category] = {
      category,
      file_name: normalizeUkdocsText(file.file_name),
      uploaded_at: normalizeUkdocsText(file.uploaded_at),
      size: Number.isFinite(Number(file.size)) ? Number(file.size) : 0,
      content_base64: String(file.content_base64 || "").trim(),
    };
  }
  return next;
}

function ukdocsUploadedFilesWithoutContent(files) {
  const normalized = normalizeUkdocsUploadedFiles(files);
  return Object.fromEntries(
    Object.entries(normalized).map(([category, file]) => [category, {
      ...file,
      content_base64: "",
    }]),
  );
}

function deriveUkdocsShipmentStatus(shipment) {
  const uploadedCategories = Object.values(shipment.uploaded_files || {}).filter((item) => item.file_name);
  if (!uploadedCategories.length) {
    return "not_started";
  }
  const requiredHeaderReady = shipment.customer_id && shipment.shipment_date && shipment.export_reference;
  if (!requiredHeaderReady) {
    return "files_uploaded";
  }
  if (shipment.audit_status === "passed") {
    return shipment.ready ? "ready" : "audit_passed";
  }
  if (shipment.audit_status === "failed") {
    return "failed";
  }
  return "validated";
}

function normalizeUkdocsShipment(shipment) {
  const uploadedFiles = normalizeUkdocsUploadedFiles(shipment?.uploaded_files);
  const categoriesIncluded = Object.values(uploadedFiles).filter((item) => item.file_name).map((item) => item.category);
  const normalized = {
    id: normalizeUkdocsText(shipment?.id) || crypto.randomUUID(),
    customer_id: normalizeUkdocsText(shipment?.customer_id),
    shipment_date: normalizeUkdocsText(shipment?.shipment_date),
    truck_number: normalizeUkdocsText(shipment?.truck_number),
    trailer_number: normalizeUkdocsText(shipment?.trailer_number),
    invoice_numbers: String(shipment?.invoice_numbers || "").trim(),
    invoice_numbers_by_category: normalizeUkdocsInvoiceNumbersByCategory(shipment?.invoice_numbers_by_category),
    export_reference: normalizeUkdocsText(shipment?.export_reference),
    currency: normalizeUkdocsText(shipment?.currency),
    delivery_terms: normalizeUkdocsText(shipment?.delivery_terms),
    uk_arrival_port: normalizeUkdocsText(shipment?.uk_arrival_port),
    transport_customs_info: String(shipment?.transport_customs_info || "").trim(),
    owner: normalizeUkdocsText(shipment?.owner),
    regulation: normalizeUkdocsText(shipment?.regulation),
    destination_country: normalizeUkdocsText(shipment?.destination_country),
    customs_office_of_exit: normalizeUkdocsText(shipment?.customs_office_of_exit),
    location: normalizeUkdocsText(shipment?.location),
    delivery_terms_city: normalizeUkdocsText(shipment?.delivery_terms_city),
    border_transport_mode: normalizeUkdocsText(shipment?.border_transport_mode),
    border_transport_nationality: normalizeUkdocsText(shipment?.border_transport_nationality),
    importer: normalizeUkdocsText(shipment?.importer),
    vessel: normalizeUkdocsText(shipment?.vessel),
    freight_costs: normalizeUkdocsText(shipment?.freight_costs),
    insurance: normalizeUkdocsText(shipment?.insurance),
    marks_and_numbers: String(shipment?.marks_and_numbers || "").trim(),
    container_number: normalizeUkdocsText(shipment?.container_number),
    uploaded_files: uploadedFiles,
    categories_included: categoriesIncluded,
    validation_warnings: Array.isArray(shipment?.validation_warnings) ? shipment.validation_warnings.map((item) => String(item || "").trim()).filter(Boolean) : [],
    audit_status: normalizeUkdocsText(shipment?.audit_status),
    ready: shipment?.ready === true,
    notes: String(shipment?.notes || "").trim(),
    print_collection_id: normalizeUkdocsText(shipment?.print_collection_id),
    reference_connect: normalizeUkdocsText(shipment?.reference_connect),
    created_by: normalizeUkdocsText(shipment?.created_by),
    created_at: normalizeUkdocsText(shipment?.created_at),
    updated_at: normalizeUkdocsText(shipment?.updated_at),
  };
  normalized.status = deriveUkdocsShipmentStatus(normalized);
  return normalized;
}

function normalizeUkdocsAuditReport(report) {
  return {
    id: normalizeUkdocsText(report?.id) || crypto.randomUUID(),
    shipment_id: normalizeUkdocsText(report?.shipment_id),
    shipment_reference: normalizeUkdocsText(report?.shipment_reference),
    shipment_date: normalizeUkdocsText(report?.shipment_date),
    customer_name: normalizeUkdocsText(report?.customer_name),
    created_at: normalizeUkdocsText(report?.created_at),
    final_status: normalizeUkdocsText(report?.final_status),
    warnings: Array.isArray(report?.warnings) ? report.warnings.map((item) => String(item || "").trim()).filter(Boolean) : [],
    summary: String(report?.summary || "").trim(),
    summary_rows: Array.isArray(report?.summary_rows) ? report.summary_rows : [],
  };
}

function normalizeUkdocsPrintDocument(document) {
  if (!document || typeof document !== "object") {
    return null;
  }
  const storageName = normalizeUkdocsText(document.storage_name);
  return storageName ? {
    storage_name: storageName,
    original_name: String(document.original_name || storageName).trim() || storageName,
    mime_type: normalizeUkdocsText(document.mime_type),
    size_bytes: Number(document.size_bytes || 0),
    saved_at: normalizeUkdocsText(document.saved_at),
    saved_by: normalizeUkdocsText(document.saved_by),
    document_kind: normalizeUkdocsText(document.document_kind),
    category: normalizeUkdocsText(document.category),
    content_type: normalizeUkdocsText(document.content_type),
    line_count: Number(document.line_count || 0),
    delimiter: normalizeUkdocsText(document.delimiter),
    parse_error: String(document.parse_error || "").trim(),
    parsed_data: document.parsed_data && typeof document.parsed_data === "object"
      ? JSON.parse(JSON.stringify(document.parsed_data))
      : null,
  } : null;
}

function ukdocsPrintDocumentIdentity(document) {
  const originalName = String(document?.original_name || "").trim().toLowerCase();
  if (originalName) {
    return originalName;
  }
  return String(document?.storage_name || "").trim().toLowerCase();
}

function normalizeUkdocsPrintDocumentList(documents) {
  if (!Array.isArray(documents)) {
    return [];
  }
  const seen = new Set();
  const normalized = [];
  for (const document of documents.map(normalizeUkdocsPrintDocument).filter(Boolean)) {
    const identity = ukdocsPrintDocumentIdentity(document);
    if (identity && seen.has(identity)) {
      continue;
    }
    if (identity) {
      seen.add(identity);
    }
    normalized.push(document);
  }
  return normalized;
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
  return (collection?.customer_id && (Array.isArray(customers) ? customers.find((item) => item.id === collection.customer_id) : null))
    || matchUkdocsCustomerForPrintCollection(customers || [], collection)
    || null;
}

function ukdocsCollectionNeedsPhyto(collection, customer = null) {
  const inspectionMode = ukdocsPrintInspectionMode(collection);
  const pdTypeCompact = String(collection?.pd_type || "").trim().toLowerCase().replace(/\s+/g, "");
  if (inspectionMode === "stock_control") {
    return false;
  }
  if (customer?.required_phyto === false) {
    return false;
  }
  if (pdTypeCompact.includes("nophytoneeded")) {
    return false;
  }
  return true;
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

function deriveUkdocsPrintCollectionStatus(collection) {
  const inspectionMode = ukdocsPrintInspectionMode(collection);
  if (inspectionMode === "stock_control") {
    const inspectionReady = Boolean(collection?.documents?.inspection_list?.storage_name);
    const locationsReady = Boolean(collection?.documents?.locations_file?.storage_name);
    if (inspectionReady && locationsReady) {
      return "complete";
    }
    if (inspectionReady || locationsReady) {
      return "partial";
    }
    return "pending";
  }
  if (inspectionMode === "reinspection") {
    const phytoReady = Array.isArray(collection?.documents?.phyto_files) && collection.documents.phyto_files.length > 0;
    const inspectionReady = Boolean(collection?.documents?.inspection_list?.storage_name);
    const extraReady = Boolean(collection?.documents?.export_extra?.storage_name);
    const generatedReady = Array.isArray(collection?.documents?.generated_files) && collection.documents.generated_files.length > 0;
    if (phytoReady && inspectionReady && (extraReady || generatedReady)) {
      return "complete";
    }
    if (phytoReady || inspectionReady || extraReady || generatedReady) {
      return "partial";
    }
    return "pending";
  }
  const phytoReady = Array.isArray(collection?.documents?.phyto_files) && collection.documents.phyto_files.length > 0;
  const extraReady = Boolean(collection?.documents?.export_extra?.storage_name);
  const generatedReady = Array.isArray(collection?.documents?.generated_files) && collection.documents.generated_files.length > 0;
  if (phytoReady && (extraReady || generatedReady)) {
    return "complete";
  }
  if (phytoReady || extraReady || generatedReady) {
    return "partial";
  }
  return "pending";
}

function normalizeUkdocsPrintCollection(collection) {
  const shipmentReference = normalizeUkdocsText(collection?.shipment_reference);
  const invoiceNumbers = String(collection?.invoice_numbers || "").trim() || shipmentReference;
  const normalized = {
    id: normalizeUkdocsText(collection?.id) || crypto.randomUUID(),
    source: normalizeUkdocsText(collection?.source) || "manual",
    shipment_id: normalizeUkdocsText(collection?.shipment_id),
    shipment_reference: shipmentReference,
    shipment_date: normalizeUkdocsText(collection?.shipment_date),
    customer_id: normalizeUkdocsText(collection?.customer_id),
    customer_name: normalizeUkdocsText(collection?.customer_name),
    collection_type: normalizeUkdocsText(collection?.collection_type) || (isHonselersdijkStockControl(collection) ? "stock_control" : "export"),
    invoice_numbers: invoiceNumbers,
    truck_number: normalizeUkdocsText(collection?.truck_number),
    trailer_number: normalizeUkdocsText(collection?.trailer_number),
    reference_connect: normalizeUkdocsText(collection?.reference_connect),
    city_name: normalizeUkdocsText(collection?.city_name),
    border_crossing: normalizeUkdocsText(collection?.border_crossing),
    hub_code: normalizeUkdocsText(collection?.hub_code),
    remark: String(collection?.remark || "").trim(),
    pd_form: normalizeUkdocsText(collection?.pd_form),
    re_export: String(collection?.re_export || "").trim(),
    pd_type: String(collection?.pd_type || "").trim(),
    pd_code: normalizeUkdocsText(collection?.pd_code),
    sheet_row_number: Number(collection?.sheet_row_number || 0),
    generated_at: normalizeUkdocsText(collection?.generated_at),
    updated_at: normalizeUkdocsText(collection?.updated_at),
    notes: String(collection?.notes || "").trim(),
    delivery_email: {
      ok: collection?.delivery_email?.ok === true,
      recipients: Array.isArray(collection?.delivery_email?.recipients) ? collection.delivery_email.recipients.map((item) => String(item || "").trim()).filter(Boolean) : [],
      sent_at: normalizeUkdocsText(collection?.delivery_email?.sent_at),
      error: String(collection?.delivery_email?.error || "").trim(),
    },
    csi_email: {
      ok: collection?.csi_email?.ok === true,
      recipients: Array.isArray(collection?.csi_email?.recipients) ? collection.csi_email.recipients.map((item) => String(item || "").trim()).filter(Boolean) : [],
      sent_at: normalizeUkdocsText(collection?.csi_email?.sent_at),
      error: String(collection?.csi_email?.error || "").trim(),
    },
    documents: {
      phyto_files: normalizeUkdocsPrintDocumentList(collection?.documents?.phyto_files || (collection?.documents?.phyto ? [collection.documents.phyto] : [])),
      export_extra: normalizeUkdocsPrintDocument(collection?.documents?.export_extra),
      generated_files: normalizeUkdocsPrintDocumentList(collection?.documents?.generated_files),
      inspection_list: normalizeUkdocsPrintDocument(collection?.documents?.inspection_list),
      locations_file: normalizeUkdocsPrintDocument(collection?.documents?.locations_file),
      temp_phyto_files: normalizeUkdocsPrintDocumentList(collection?.documents?.temp_phyto_files || (collection?.documents?.temp_phyto ? [collection.documents.temp_phyto] : [])),
      temp_phyto_plants_file: normalizeUkdocsPrintDocument(collection?.documents?.temp_phyto_plants_file),
      temp_phyto_plants_xml_file: normalizeUkdocsPrintDocument(collection?.documents?.temp_phyto_plants_xml_file),
      ipaffs_file: normalizeUkdocsPrintDocument(collection?.documents?.ipaffs_file),
      ipaffs_plants_file: normalizeUkdocsPrintDocument(collection?.documents?.ipaffs_plants_file),
    },
    csi_report: normalizeUkdocsCsiReport(collection?.csi_report),
  };
  normalized.status = deriveUkdocsPrintCollectionStatus(normalized);
  return normalized;
}

function ukdocsPrintSplitTokens(value) {
  return String(value || "")
    .split(/[\/,\s;]+/)
    .map((item) => String(item || "").trim())
    .filter(Boolean);
}

function getUkdocsPrintCollectionRequirements(collection, customers) {
  const inspectionMode = ukdocsPrintInspectionMode(collection);
  if (inspectionMode === "stock_control") {
    const missing = [];
    if (!collection?.documents?.inspection_list?.storage_name) {
      missing.push("Inspection list");
    }
    if (!collection?.documents?.locations_file?.storage_name) {
      missing.push("Locations file");
    }
    return {
      customer: null,
      missing,
      complete: missing.length === 0,
    };
  }
  if (inspectionMode === "reinspection") {
    const customer = (collection?.customer_id && (Array.isArray(customers) ? customers.find((item) => item.id === collection.customer_id) : null))
      || matchUkdocsCustomerForPrintCollection(customers || [], collection)
      || null;
    const phytoCount = (collection?.documents?.phyto_files || []).length;
    const phytoExpected = ukdocsPrintSplitTokens(collection?.reference_connect).length;
    const generatedFiles = collection?.documents?.generated_files || [];
    const generatedExportReady = generatedFiles.some((file) => file.document_kind === "export");
    const generatedInvoiceCount = countUkdocsGeneratedInvoiceGroups(generatedFiles);
    const invoiceExpected = ukdocsPrintSplitTokens(collection?.invoice_numbers).length;
    const missing = [];
    if (ukdocsCollectionNeedsPhyto(collection, customer) && phytoExpected > 0 && phytoCount < phytoExpected) {
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
      complete: missing.length === 0,
    };
  }
  const customer = (collection?.customer_id && (Array.isArray(customers) ? customers.find((item) => item.id === collection.customer_id) : null))
    || matchUkdocsCustomerForPrintCollection(customers || [], collection)
    || null;
  const phytoCount = (collection?.documents?.phyto_files || []).length;
  const phytoExpected = ukdocsPrintSplitTokens(collection?.reference_connect).length;
  const generatedFiles = collection?.documents?.generated_files || [];
  const generatedExportReady = generatedFiles.some((file) => file.document_kind === "export");
  const generatedInvoiceCount = countUkdocsGeneratedInvoiceGroups(generatedFiles);
  const invoiceExpected = ukdocsPrintSplitTokens(collection?.invoice_numbers).length;
  const missing = [];
  if (ukdocsCollectionNeedsPhyto(collection, customer) && phytoExpected > 0 && phytoCount < phytoExpected) {
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
  return {
    customer,
    missing,
    complete: missing.length === 0,
  };
}

function summarizeUkdocsWarnings(warnings) {
  if (!Array.isArray(warnings) || !warnings.length) {
    return "";
  }
  return warnings
    .map((warning) => String(warning?.message || warning || "").trim())
    .filter(Boolean)
    .join(" | ");
}

function normalizeUkdocsState(state) {
  return {
    company_settings: normalizeUkdocsCompanySettings(state?.company_settings || defaultUkdocsState.company_settings),
    export_defaults: normalizeUkdocsExportDefaults(state?.export_defaults || defaultUkdocsState.export_defaults),
    templates: normalizeUkdocsTemplates(state?.templates || defaultUkdocsState.templates),
    customers: Array.isArray(state?.customers) ? state.customers.map(normalizeUkdocsCustomer).sort((a, b) => a.customer_name.localeCompare(b.customer_name)) : [],
    column_mappings: normalizeUkdocsColumnMappings(state?.column_mappings || defaultUkdocsState.column_mappings),
    shipments: Array.isArray(state?.shipments) ? state.shipments.map(normalizeUkdocsShipment).sort((a, b) => String(b.shipment_date || b.updated_at).localeCompare(String(a.shipment_date || a.updated_at))) : [],
    audit_reports: Array.isArray(state?.audit_reports) ? state.audit_reports.map(normalizeUkdocsAuditReport).sort((a, b) => String(b.created_at).localeCompare(String(a.created_at))) : [],
    print_collections: Array.isArray(state?.print_collections) ? state.print_collections.map(normalizeUkdocsPrintCollection).sort((a, b) => String(b.shipment_date || b.updated_at).localeCompare(String(a.shipment_date || a.updated_at))) : [],
  };
}

function ukdocsViewerSettingsSummary(settings) {
  return {
    gmail_connected_email: String(settings?.gmail_connected_email || "").trim(),
    ukdocs_print_spreadsheet_id: String(settings?.ukdocs_print_spreadsheet_id || "").trim(),
  };
}

async function readUkdocsState() {
  const payload = await readJsonFile(ukdocsStatePath, defaultUkdocsState);
  return normalizeUkdocsState(payload);
}

async function writeUkdocsState(state) {
  await writeJsonFile(ukdocsStatePath, normalizeUkdocsState(state));
}

function normalizeExpeditionStickerFile(file) {
  if (!file || typeof file !== "object") {
    return null;
  }
  const storageName = String(file.storage_name || "").trim();
  return storageName ? {
    storage_name: storageName,
    original_name: String(file.original_name || storageName).trim() || storageName,
    saved_at: String(file.saved_at || "").trim(),
    saved_by: String(file.saved_by || "").trim(),
    size_bytes: Number(file.size_bytes || 0),
  } : null;
}

function normalizeExpeditionStickerState(state) {
  return {
    planning_file: normalizeExpeditionStickerFile(state?.planning_file),
    split_file: normalizeExpeditionStickerFile(state?.split_file),
  };
}

async function readExpeditionStickerState() {
  const payload = await readJsonFile(expeditionStickerStatePath, defaultExpeditionStickerState);
  return normalizeExpeditionStickerState(payload);
}

async function writeExpeditionStickerState(state) {
  await writeJsonFile(expeditionStickerStatePath, normalizeExpeditionStickerState(state));
}

function expeditionStickerFilePath(file) {
  if (!file?.storage_name) {
    return "";
  }
  return path.join(expeditionStickerFilesDir, file.storage_name);
}

async function inspectExpeditionStickerSource(kind, filePath) {
  const output = await runExpeditionStickerWorker([
    "inspect-source",
    "--kind",
    kind,
    "--input",
    filePath,
  ]);
  return JSON.parse(output.toString("utf8"));
}

async function saveExpeditionStickerUpload(kind, filePayload, requestUser) {
  const originalName = path.basename(String(filePayload?.name || "").trim());
  const contentBase64 = String(filePayload?.content_base64 || "").trim();
  if (!originalName || !contentBase64) {
    throw new Error(`Choose a ${kind} file first`);
  }

  const extension = path.extname(originalName).toLowerCase() || ".xlsx";
  const storageName = `${kind}${extension}`;
  const fileBuffer = Buffer.from(contentBase64, "base64");
  await fs.mkdir(expeditionStickerFilesDir, { recursive: true });
  await fs.writeFile(path.join(expeditionStickerFilesDir, storageName), fileBuffer);

  return {
    storage_name: storageName,
    original_name: originalName,
    saved_at: new Date().toISOString(),
    saved_by: requestUser.username,
    size_bytes: fileBuffer.length,
  };
}

async function deleteExpeditionStickerUpload(kind) {
  const normalizedKind = kind === "split" ? "split" : kind === "planning" ? "planning" : "";
  if (!normalizedKind) {
    throw new Error("Unknown expedition sticker source");
  }

  const currentState = await readExpeditionStickerState();
  const file = currentState[`${normalizedKind}_file`];
  if (file?.storage_name) {
    await fs.rm(path.join(expeditionStickerFilesDir, file.storage_name), { force: true }).catch(() => {});
  }

  const nextState = {
    ...currentState,
    [`${normalizedKind}_file`]: null,
  };
  await writeExpeditionStickerState(nextState);
  return nextState;
}

function emptyFustMetrics() {
  return {
    dc: 0,
    cctag: 0,
    dcs: 0,
    dco: 0,
    pal: 0,
    vk: 0,
  };
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
    fustbon_reference: String(action?.fustbon_reference || "").trim(),
    fustfactuur_reference: String(action?.fustfactuur_reference || "").trim(),
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
    confirmed_at: String(action?.confirmed_at || ""),
    confirmed_by: String(action?.confirmed_by || ""),
    import_source: normalizeFustImportSource(action?.import_source),
    deleted: action?.deleted === true,
    deleted_at: String(action?.deleted_at || ""),
    deleted_by: String(action?.deleted_by || ""),
    sheet_sync: action?.sheet_sync || { ok: false, target_sheets: [], error: "Not attempted" },
    email_sync: action?.email_sync || { ok: false, recipients: [], error: "Not attempted" },
    confirmation_reminder: normalizeFustConfirmationReminder(action?.confirmation_reminder),
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

async function mirrorFustActionToDatabase(action) {
  if (!isDatabaseEnabled()) {
    return;
  }
  try {
    await saveFustActionToDatabase(normalizeFustAction(action));
  } catch (error) {
    console.error("Fust database mirror failed:", error instanceof Error ? error.message : String(error));
  }
}

async function mirrorFustDeleteToDatabase(actionId) {
  if (!isDatabaseEnabled()) {
    return;
  }
  try {
    await markFustActionDeletedInDatabase(String(actionId || "").trim());
  } catch (error) {
    console.error("Fust database delete mirror failed:", error instanceof Error ? error.message : String(error));
  }
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

async function readDagFoutjesState() {
  const payload = await readJsonFile(dagFoutjesStatePath, defaultDagFoutjesState);
  const shared = payload && typeof payload.shared === "object" && !Array.isArray(payload.shared)
    ? payload.shared
    : {};
  return { shared };
}

async function writeDagFoutjesState(state) {
  await writeJsonFile(dagFoutjesStatePath, {
    shared: state && typeof state.shared === "object" && !Array.isArray(state.shared) ? state.shared : {},
  });
}

function dagFoutjesIsoWeekParts(dateStr) {
  const normalized = String(dateStr || "").slice(0, 10);
  const parsed = new Date(`${normalized}T00:00:00`);
  if (Number.isNaN(parsed.getTime())) {
    return { iso_week: "", iso_week_label: "", month: "" };
  }
  const utcDate = new Date(Date.UTC(parsed.getFullYear(), parsed.getMonth(), parsed.getDate()));
  const dayNumber = utcDate.getUTCDay() || 7;
  utcDate.setUTCDate(utcDate.getUTCDate() + 4 - dayNumber);
  const isoYear = utcDate.getUTCFullYear();
  const yearStart = new Date(Date.UTC(isoYear, 0, 1));
  const week = Math.ceil((((utcDate - yearStart) / 86400000) + 1) / 7);
  const isoWeek = `${isoYear}-W${String(week).padStart(2, "0")}`;
  const month = normalized.slice(0, 7);
  return { iso_week: isoWeek, iso_week_label: isoWeek, month };
}

function normalizeDagFoutjesOverviewEntry(entry, dateKey) {
  const normalizedDate = String(dateKey || entry?.dateKey || "").slice(0, 10);
  const parts = dagFoutjesIsoWeekParts(normalizedDate);
  const typeKey = String(entry?.type || "").trim();
  return {
    id: String(entry?.id || ""),
    date: normalizedDate,
    iso_week: parts.iso_week,
    month: parts.month,
    person_id: String(entry?.personId || "").trim(),
    person_name: String(entry?.personName || "").trim(),
    type_key: typeKey,
    type_label: dagFoutjesEnglishTypeLabel(typeKey, entry?.typeLabel || entry?.type || ""),
    comment: String(entry?.comment || ""),
    time: String(entry?.time || "").trim(),
    timestamp: Number(entry?.timestamp || 0) || 0,
  };
}

async function buildDagFoutjesOverviewPayload() {
  const state = await readDagFoutjesState();
  const keys = Object.keys(state.shared || {}).filter((key) => key.startsWith("errors:")).sort();
  const entries = [];

  for (const key of keys) {
    const dateKey = key.replace(/^errors:/, "").slice(0, 10);
    if (!dateKey) {
      continue;
    }
    let dayEntries = [];
    try {
      dayEntries = JSON.parse(String(state.shared[key] || "[]"));
    } catch {
      dayEntries = [];
    }
    if (!Array.isArray(dayEntries)) {
      continue;
    }
    for (const entry of dayEntries) {
      const normalized = normalizeDagFoutjesOverviewEntry(entry, dateKey);
      if (!normalized.person_name || !normalized.type_key || !normalized.date) {
        continue;
      }
      entries.push(normalized);
    }
  }

  entries.sort((left, right) => {
    if (left.date !== right.date) {
      return left.date.localeCompare(right.date);
    }
    return left.timestamp - right.timestamp;
  });

  const byType = new Map();
  const byPerson = new Map();
  const byDay = new Map();
  const byWeek = new Map();
  const byMonth = new Map();

  const bumpTypeMap = (map, key, label) => {
    const current = map.get(key) || { key, label, count: 0 };
    current.count += 1;
    map.set(key, current);
  };

  for (const entry of entries) {
    bumpTypeMap(byType, entry.type_key, entry.type_label);

    const person = byPerson.get(entry.person_id) || {
      person_id: entry.person_id,
      person_name: entry.person_name,
      total: 0,
      types: {},
    };
    person.total += 1;
    person.types[entry.type_label] = (person.types[entry.type_label] || 0) + 1;
    byPerson.set(entry.person_id, person);

    const day = byDay.get(entry.date) || { period: entry.date, total: 0, people: new Set(), types: {} };
    day.total += 1;
    day.people.add(entry.person_name);
    day.types[entry.type_label] = (day.types[entry.type_label] || 0) + 1;
    byDay.set(entry.date, day);

    const week = byWeek.get(entry.iso_week) || { period: entry.iso_week, total: 0, people: new Set(), types: {} };
    week.total += 1;
    week.people.add(entry.person_name);
    week.types[entry.type_label] = (week.types[entry.type_label] || 0) + 1;
    byWeek.set(entry.iso_week, week);

    const month = byMonth.get(entry.month) || { period: entry.month, total: 0, people: new Set(), types: {} };
    month.total += 1;
    month.people.add(entry.person_name);
    month.types[entry.type_label] = (month.types[entry.type_label] || 0) + 1;
    byMonth.set(entry.month, month);
  }

  const summarizePeriodMap = (map) =>
    [...map.values()]
      .map((item) => ({
        period: item.period,
        total: item.total,
        people_count: item.people.size,
        top_type: Object.entries(item.types).sort((a, b) => b[1] - a[1])[0]?.[0] || "",
        types: item.types,
      }))
      .sort((left, right) => right.period.localeCompare(left.period));

  return {
    total_entries: entries.length,
    people_count: byPerson.size,
    type_totals: [...byType.values()].sort((left, right) => right.count - left.count || left.label.localeCompare(right.label)),
    person_totals: [...byPerson.values()].sort((left, right) => right.total - left.total || left.person_name.localeCompare(right.person_name)),
    day_totals: summarizePeriodMap(byDay),
    week_totals: summarizePeriodMap(byWeek),
    month_totals: summarizePeriodMap(byMonth),
    entries,
  };
}

function resolveDagFoutjesHtmlPath() {
  return dagFoutjesHtmlPathCandidates.find((candidate) => existsSync(candidate)) || null;
}

async function mergeDagFoutjesPeopleFromClock(state) {
  const nextState = {
    shared: state && typeof state.shared === "object" && !Array.isArray(state.shared) ? { ...state.shared } : {},
  };
  const settings = await readFustSettings();
  if (!settings.clock_spreadsheet_id || !settings.clock_employee_sheet_name) {
    return { state: nextState, changed: false };
  }

  let existingPeople = [];
  try {
    existingPeople = JSON.parse(String(nextState.shared.people || "[]"));
  } catch {
    existingPeople = [];
  }
  if (!Array.isArray(existingPeople)) {
    existingPeople = [];
  }

  const rows = await loadSheetRows(settings.clock_spreadsheet_id, settings.clock_employee_sheet_name);
  const employees = buildClockEmployeesFromSheetRows(rows).employees || [];
  const people = existingPeople
    .filter((person) => person && typeof person === "object")
    .map((person) => ({
      id: String(person.id || "").trim(),
      name: String(person.name || "").trim(),
    }))
    .filter((person) => person.id && person.name);

  const byId = new Map(people.map((person) => [person.id, person]));
  const byName = new Map(people.map((person) => [person.name.toLowerCase(), person]));
  let changed = false;

  for (const employee of employees) {
    const employeeName = String(employee?.name || "").trim();
    const employeeId = `clock-${String(employee?.tbnr || "").trim().toUpperCase()}`;
    if (!employeeName || employeeId === "clock-") {
      continue;
    }
    const existing = byId.get(employeeId) || byName.get(employeeName.toLowerCase()) || null;
    if (existing) {
      if (!byId.has(employeeId)) {
        existing.id = employeeId;
        byId.set(employeeId, existing);
        changed = true;
      }
      if (existing.name !== employeeName) {
        byName.delete(existing.name.toLowerCase());
        existing.name = employeeName;
        byName.set(employeeName.toLowerCase(), existing);
        changed = true;
      }
      continue;
    }
    const person = { id: employeeId, name: employeeName };
    people.push(person);
    byId.set(employeeId, person);
    byName.set(employeeName.toLowerCase(), person);
    changed = true;
  }

  if (changed) {
    nextState.shared.people = JSON.stringify(
      people.sort((left, right) => left.name.localeCompare(right.name)),
    );
  }

  return { state: nextState, changed };
}

const dagFoutjesSheetHeaders = ["week", "day", "name", "type", "comment"];
const dagFoutjesEnglishTypeLabels = {
  za_duzo: "too much",
  za_malo: "too little",
  zla_polka: "wrong shelf",
  zle_kwiaty: "wrong flowers",
  inne: "other",
};

function dagFoutjesEnglishTypeLabel(typeKey, fallbackLabel = "") {
  const normalizedKey = String(typeKey || "").trim();
  return dagFoutjesEnglishTypeLabels[normalizedKey] || String(fallbackLabel || normalizedKey).trim();
}

function dagFoutjesDateParts(dateStr) {
  const normalized = String(dateStr || "").slice(0, 10);
  const parsed = new Date(`${normalized}T00:00:00`);
  if (Number.isNaN(parsed.getTime())) {
    return { date: normalized, week: "", day: normalized };
  }
  const day = normalized;
  const utcDate = new Date(Date.UTC(parsed.getFullYear(), parsed.getMonth(), parsed.getDate()));
  const dayNumber = utcDate.getUTCDay() || 7;
  utcDate.setUTCDate(utcDate.getUTCDate() + 4 - dayNumber);
  const yearStart = new Date(Date.UTC(utcDate.getUTCFullYear(), 0, 1));
  const week = Math.ceil((((utcDate - yearStart) / 86400000) + 1) / 7);
  return { date: normalized, week: String(week), day };
}

function normalizeDagFoutjesSheetSync(value) {
  return {
    row_number: Math.max(0, Number(value?.row_number || 0)),
    synced_at: String(value?.synced_at || ""),
    ok: value?.ok !== false,
    error: String(value?.error || ""),
  };
}

function normalizeDagFoutjesEntry(entry, fallbackDate = "") {
  const typeKey = String(entry?.type || "").trim();
  return {
    id: String(entry?.id || crypto.randomUUID()),
    personId: String(entry?.personId || "").trim(),
    personName: String(entry?.personName || "").trim(),
    type: typeKey,
    typeLabel: dagFoutjesEnglishTypeLabel(typeKey, entry?.typeLabel || entry?.type || ""),
    comment: String(entry?.comment || ""),
    time: String(entry?.time || "").trim(),
    timestamp: Number(entry?.timestamp || 0) || Date.now(),
    dateKey: String(entry?.dateKey || fallbackDate || "").slice(0, 10),
    sheet_sync: normalizeDagFoutjesSheetSync(entry?.sheet_sync || {}),
  };
}

async function ensureDagFoutjesSheetHeader(spreadsheetId, sheetName) {
  const rows = await loadSheetRows(spreadsheetId, sheetName);
  const firstRow = Array.isArray(rows[0]) ? rows[0].map((value) => String(value || "").trim().toLowerCase()) : [];
  const matchesHeader = dagFoutjesSheetHeaders.every((value, index) => firstRow[index] === value);
  if (!matchesHeader) {
    await writeSheetRowAt(spreadsheetId, sheetName, 1, dagFoutjesSheetHeaders);
  }
}

function buildDagFoutjesSheetRow(dateStr, entry) {
  const parts = dagFoutjesDateParts(dateStr || entry?.dateKey || "");
  return [
    parts.week,
    parts.day,
    String(entry?.personName || "").trim(),
    String(entry?.typeLabel || entry?.type || "").trim(),
    String(entry?.comment || ""),
  ];
}

async function syncDagFoutjesEntriesToSheet(entries, dateStr, settings) {
  if (!settings.clock_spreadsheet_id) {
    throw new Error("Clock spreadsheet ID is not configured");
  }
  const sheetName = "fouten";
  await ensureDagFoutjesSheetHeader(settings.clock_spreadsheet_id, sheetName);

  const normalizedDate = String(dateStr || "").slice(0, 10);
  const nextEntries = [];
  for (const rawEntry of Array.isArray(entries) ? entries : []) {
    const entry = normalizeDagFoutjesEntry(rawEntry, normalizedDate);
    const rowPayload = buildDagFoutjesSheetRow(normalizedDate, entry);
    const rowNumber = Math.max(0, Number(entry?.sheet_sync?.row_number || 0));
    const output = rowNumber >= 2
      ? await writeSheetRowAt(settings.clock_spreadsheet_id, sheetName, rowNumber, rowPayload)
      : await writeSheetRowToFirstEmpty(settings.clock_spreadsheet_id, sheetName, rowPayload);
    nextEntries.push({
      ...entry,
      sheet_sync: {
        ok: true,
        error: "",
        synced_at: new Date().toISOString(),
        row_number: Number(output?.row_number || rowNumber || 0),
      },
    });
  }
  return nextEntries;
}

async function clearDagFoutjesEntryFromSheet(entry, settings) {
  if (!settings.clock_spreadsheet_id) {
    throw new Error("Clock spreadsheet ID is not configured");
  }
  const rowNumber = Math.max(0, Number(entry?.sheet_sync?.row_number || 0));
  if (rowNumber < 2) {
    return { ok: true, row_number: 0 };
  }
  await writeSheetRowAt(settings.clock_spreadsheet_id, "fouten", rowNumber, emptySheetRow(dagFoutjesSheetHeaders.length));
  return { ok: true, row_number: rowNumber };
}

let dagFoutjesSheetSyncRunning = false;

async function syncPendingDagFoutjesDaysToSheet() {
  if (dagFoutjesSheetSyncRunning) {
    return;
  }
  dagFoutjesSheetSyncRunning = true;
  try {
    const settings = await readFustSettings();
    if (!settings.clock_spreadsheet_id) {
      return;
    }
    const state = await readDagFoutjesState();
    const today = localDateIso();
    const keys = Object.keys(state.shared || {}).filter((key) => key.startsWith("errors:")).sort();
    let changed = false;
    for (const key of keys) {
      const dateKey = key.replace(/^errors:/, "").slice(0, 10);
      if (!dateKey || dateKey >= today) {
        continue;
      }
      let entries = [];
      try {
        entries = JSON.parse(String(state.shared[key] || "[]"));
      } catch {
        entries = [];
      }
      if (!Array.isArray(entries) || !entries.length) {
        continue;
      }
      const needsSync = entries.some((entry) => Number(entry?.sheet_sync?.row_number || 0) < 2);
      if (!needsSync) {
        continue;
      }
      const syncedEntries = await syncDagFoutjesEntriesToSheet(entries, dateKey, settings);
      state.shared[key] = JSON.stringify(syncedEntries);
      changed = true;
    }
    if (changed) {
      await writeDagFoutjesState(state);
    }
  } catch (error) {
    console.error("Dag Foutjes sheet auto-sync failed:", error instanceof Error ? error.message : String(error));
  } finally {
    dagFoutjesSheetSyncRunning = false;
  }
}

function dagFoutjesBridgeScript() {
  return `<script>
window.storage = {
  async get(key, shared) {
    if (!shared) {
      const value = window.localStorage.getItem(String(key || ""));
      return value === null ? null : { key: String(key || ""), value };
    }
    const response = await fetch('/api/dag-foutjes/storage?op=get&key=' + encodeURIComponent(String(key || '')), {
      credentials: 'same-origin',
      cache: 'no-store',
    });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(payload.error || ('Request failed with ' + response.status));
    }
    return payload.found ? { key: payload.key, value: payload.value } : null;
  },
  async set(key, value, shared) {
    if (!shared) {
      window.localStorage.setItem(String(key || ''), String(value ?? ''));
      return { ok: true, key: String(key || '') };
    }
    const response = await fetch('/api/dag-foutjes/storage', {
      method: 'POST',
      credentials: 'same-origin',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ op: 'set', key: String(key || ''), value: String(value ?? '') }),
    });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(payload.error || ('Request failed with ' + response.status));
    }
    return payload;
  },
  async list(prefix, shared) {
    if (!shared) {
      const keys = [];
      const normalizedPrefix = String(prefix || '');
      for (let index = 0; index < window.localStorage.length; index += 1) {
        const key = window.localStorage.key(index);
        if (key && key.startsWith(normalizedPrefix)) {
          keys.push(key);
        }
      }
      return { keys };
    }
    const response = await fetch('/api/dag-foutjes/storage?op=list&prefix=' + encodeURIComponent(String(prefix || '')), {
      credentials: 'same-origin',
      cache: 'no-store',
    });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(payload.error || ('Request failed with ' + response.status));
    }
    return { keys: Array.isArray(payload.keys) ? payload.keys : [] };
  }
};
</script>`;
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

function clockBackupHeaders() {
  return ["Date", "TBNR", "Name", "Type", "IN", "OUT", "Worked time", "Source", "ID", "Created by", "Created at", "Status"];
}

function clockBackupRow(sessionRow, status = "active") {
  return [...sessionRow.slice(0, 11), status];
}

function applyClockSessionSync(records, session, rowNumber, sheetName) {
  if (rowNumber < 2) {
    return;
  }
  const syncInfo = {
    ok: true,
    target_sheets: [sheetName],
    error: "",
    synced_at: new Date().toISOString(),
    row_number: rowNumber,
  };
  for (const recordId of [session.inRecord?.id, session.outRecord?.id]) {
    if (!recordId) {
      continue;
    }
    const recordIndex = records.findIndex((item) => item.id === recordId);
    if (recordIndex >= 0) {
      records[recordIndex] = {
        ...records[recordIndex],
        sheet_sync: {
          ...records[recordIndex].sheet_sync,
          ...syncInfo,
        },
      };
    }
  }
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
  const backupRow = clockBackupRow(sessionRow, "active");
  const output = record.direction === "OUT" && previousRowNumber >= 2
    ? await writeSheetRowAt(settings.clock_spreadsheet_id, settings.clock_records_sheet_name, previousRowNumber, backupRow)
    : await writeSheetRowToFirstEmpty(settings.clock_spreadsheet_id, settings.clock_records_sheet_name, backupRow);

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

function clockBackupEntriesForDate(rows, targetDate) {
  if (!Array.isArray(rows) || rows.length < 2) {
    return [];
  }

  const headers = rows[0].map(normalizeHeader);
  const dateIndex = firstMatchingIndex(headers, ["date", "datum"]);
  const idIndex = firstMatchingIndex(headers, ["id"]);
  const statusIndex = firstMatchingIndex(headers, ["status"]);
  return rows
    .slice(1)
    .map((row, offset) => ({
      row_number: offset + 2,
      row,
      action_date: parseSheetDateToIso(rowValue(row, dateIndex >= 0 ? dateIndex : 0)).slice(0, 10),
      id: rowValue(row, idIndex >= 0 ? idIndex : 8),
      status: rowValue(row, statusIndex >= 0 ? statusIndex : 11).toLowerCase() || "active",
    }))
    .filter((entry) => entry.action_date === targetDate)
    .sort((left, right) => left.row_number - right.row_number);
}

async function ensureClockBackupHeader(settings, rows) {
  const existingHeaders = Array.isArray(rows[0]) ? rows[0].map(normalizeHeader) : [];
  if (existingHeaders.includes("status")) {
    return;
  }
  await writeSheetRowAt(settings.clock_spreadsheet_id, settings.clock_records_sheet_name, 1, clockBackupHeaders());
}

async function rewriteClockBackupDate(date, records, settings) {
  if (!settings.clock_spreadsheet_id || !settings.clock_records_sheet_name) {
    return;
  }

  const rows = await loadSheetRows(settings.clock_spreadsheet_id, settings.clock_records_sheet_name);
  await ensureClockBackupHeader(settings, rows);
  const entries = clockBackupEntriesForDate(rows, date);
  const dateRecords = records.filter((record) => record.action_date === date);
  const sessions = clockSessionRows(dateRecords);
  const sessionsById = new Map(sessions.map((session) => [String(session.row[8] || ""), session]));
  const matchedIds = new Set();

  for (const entry of entries) {
    const matchingSession = entry.id ? sessionsById.get(entry.id) : null;
    if (matchingSession) {
      await writeSheetRowAt(
        settings.clock_spreadsheet_id,
        settings.clock_records_sheet_name,
        entry.row_number,
        clockBackupRow(matchingSession.row, "active"),
      );
      applyClockSessionSync(records, matchingSession, entry.row_number, settings.clock_records_sheet_name);
      matchedIds.add(entry.id);
      continue;
    }

    if (entry.status !== "deleted") {
      await writeSheetRowAt(
        settings.clock_spreadsheet_id,
        settings.clock_records_sheet_name,
        entry.row_number,
        clockBackupRow(entry.row, "deleted"),
      );
    }
  }

  for (const session of sessions) {
    const sessionId = String(session.row[8] || "");
    if (sessionId && matchedIds.has(sessionId)) {
      continue;
    }
    const output = await writeSheetRowToFirstEmpty(
      settings.clock_spreadsheet_id,
      settings.clock_records_sheet_name,
      clockBackupRow(session.row, "active"),
    );
    applyClockSessionSync(records, session, Number(output?.row_number || 0), settings.clock_records_sheet_name);
  }
}

async function rewriteClockBackupDates(dates, records, settings) {
  const uniqueDates = [...new Set(dates.map((value) => String(value || "").slice(0, 10)).filter(Boolean))];
  for (const date of uniqueDates) {
    await rewriteClockBackupDate(date, records, settings);
  }
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
  const statusIndex = firstMatchingIndex(headers, ["status"]);

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
    const status = rowValue(row, statusIndex >= 0 ? statusIndex : 11).toLowerCase() || "active";

    if (!tbnr || !name || status === "deleted") {
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

async function readFustBackupSnapshot(filename) {
  const safeName = path.basename(String(filename || "").trim());
  if (!safeName) {
    throw new Error("Choose a backup file first");
  }
  const resolvedPath = path.resolve(fustBackupDir, safeName);
  if (!resolvedPath.startsWith(path.resolve(fustBackupDir))) {
    throw new Error("Forbidden backup path");
  }
  if (!existsSync(resolvedPath)) {
    throw new Error("Backup not found");
  }
  const payload = await readJsonFile(resolvedPath, { actions: [] });
  return {
    filename: safeName,
    actions: Array.isArray(payload?.actions) ? payload.actions.map(normalizeFustAction) : [],
  };
}

function fustDocumentHasSavedValue(documentInfo) {
  return ["uploaded", "skipped", "failed"].includes(String(documentInfo?.status || "")) && (
    String(documentInfo?.file_id || "").trim()
    || String(documentInfo?.file_name || "").trim()
    || String(documentInfo?.web_link || "").trim()
    || String(documentInfo?.uploaded_at || "").trim()
    || String(documentInfo?.error || "").trim()
    || String(documentInfo?.status || "").trim() === "skipped"
  );
}

function mergeMissingFustActionData(currentAction, backupAction) {
  const mergedAction = normalizeFustAction({
    ...currentAction,
    fustbon_reference: currentAction?.fustbon_reference || backupAction?.fustbon_reference || "",
    fustfactuur_reference: currentAction?.fustfactuur_reference || backupAction?.fustfactuur_reference || "",
    cmr: fustDocumentHasSavedValue(currentAction?.cmr) ? currentAction?.cmr : backupAction?.cmr,
    fustbon: fustDocumentHasSavedValue(currentAction?.fustbon) ? currentAction?.fustbon : backupAction?.fustbon,
  });

  const changes = {
    fustbon_reference_restored: !currentAction?.fustbon_reference && !!mergedAction.fustbon_reference,
    fustfactuur_reference_restored: !currentAction?.fustfactuur_reference && !!mergedAction.fustfactuur_reference,
    cmr_restored: !fustDocumentHasSavedValue(currentAction?.cmr) && fustDocumentHasSavedValue(mergedAction?.cmr),
    fustbon_restored: !fustDocumentHasSavedValue(currentAction?.fustbon) && fustDocumentHasSavedValue(mergedAction?.fustbon),
  };

  return { mergedAction, changes };
}

async function restoreMissingFustDataFromBackup(filename, requestUser) {
  const backup = await readFustBackupSnapshot(filename);
  const settings = await readFustSettings();
  const [localActions, retourRows, uitgaandRows] = await Promise.all([
    readFustActions(),
    loadSheetRows(settings.spreadsheet_id, settings.in_sheet_name).catch(() => []),
    loadSheetRows(settings.spreadsheet_id, settings.out_sheet_name).catch(() => []),
  ]);

  const sheetActions = [
    ...parseRegistrySheetRows(retourRows, "IN"),
    ...parseRegistrySheetRows(uitgaandRows, "OUT"),
  ];
  const currentCandidates = [...localActions, ...sheetActions];
  const currentById = new Map(currentCandidates.map((action) => [String(action.id || "").trim(), action]).filter(([id]) => id));
  const currentBySignature = new Map(currentCandidates.map((action) => [buildActionSignature(action), action]));

  const nextLocalMap = new Map(localActions.map((action) => [buildActionMergeKey(action), action]));
  const restoredActions = [];
  const summary = {
    matched: 0,
    updated: 0,
    cmr_restored: 0,
    fustbon_restored: 0,
    fustbon_reference_restored: 0,
    fustfactuur_reference_restored: 0,
    skipped: 0,
  };

  for (const backupAction of backup.actions) {
    if (backupAction.deleted) {
      summary.skipped += 1;
      continue;
    }
    const currentMatch = currentById.get(String(backupAction.id || "").trim()) || currentBySignature.get(buildActionSignature(backupAction)) || null;
    if (!currentMatch) {
      summary.skipped += 1;
      continue;
    }
    summary.matched += 1;
    const existingLocalKey = buildActionMergeKey(currentMatch);
    const baseAction = nextLocalMap.get(existingLocalKey) || currentMatch;
    const { mergedAction, changes } = mergeMissingFustActionData(baseAction, backupAction);
    const changed = changes.cmr_restored || changes.fustbon_restored || changes.fustbon_reference_restored || changes.fustfactuur_reference_restored;
    if (!changed) {
      continue;
    }
    mergedAction.created_by = baseAction.created_by || backupAction.created_by || requestUser.username;
    mergedAction.created_at = baseAction.created_at || backupAction.created_at || new Date().toISOString();
    nextLocalMap.set(existingLocalKey, mergedAction);
    restoredActions.push(mergedAction);
    summary.updated += 1;
    if (changes.cmr_restored) {
      summary.cmr_restored += 1;
    }
    if (changes.fustbon_restored) {
      summary.fustbon_restored += 1;
    }
    if (changes.fustbon_reference_restored) {
      summary.fustbon_reference_restored += 1;
    }
    if (changes.fustfactuur_reference_restored) {
      summary.fustfactuur_reference_restored += 1;
    }
  }

  const nextLocalActions = [...nextLocalMap.values()];
  await writeFustActions(nextLocalActions);
  for (const restoredAction of restoredActions) {
    await mirrorFustActionToDatabase(restoredAction);
  }
  return {
    filename: backup.filename,
    summary,
    action_count: nextLocalActions.length,
  };
}

async function collectCurrentFustActions(settings) {
  const localActions = await readFustActions();
  let inSheetActions = [];
  let outSheetActions = [];

  try {
    const retourRows = await loadSheetRows(settings.spreadsheet_id, settings.in_sheet_name);
    inSheetActions = parseRegistrySheetRows(retourRows, "IN");
  } catch {
    inSheetActions = [];
  }

  try {
    const uitgaandRows = await loadSheetRows(settings.spreadsheet_id, settings.out_sheet_name);
    outSheetActions = parseRegistrySheetRows(uitgaandRows, "OUT");
  } catch {
    outSheetActions = [];
  }

  const deletedActionIds = new Set(
    localActions
      .filter((action) => action.deleted)
      .map((action) => String(action.id || "").trim())
      .filter(Boolean),
  );

  const dedupedActions = new Map();
  for (const action of [...inSheetActions, ...outSheetActions, ...localActions]) {
    const actionId = String(action?.id || "").trim();
    if (actionId && deletedActionIds.has(actionId)) {
      continue;
    }
    if (action.deleted) {
      continue;
    }
    dedupedActions.set(buildActionMergeKey(action), normalizeFustAction(action));
  }

  return {
    activeActions: [...dedupedActions.values()],
    deletedActionIds: [...deletedActionIds],
    localActionCount: localActions.length,
    sheetActionCount: inSheetActions.length + outSheetActions.length,
  };
}

async function applyFustImportRows(filePayload, requestUser, selectedImportKeys = []) {
  const prepared = await prepareFustImportRows(filePayload, requestUser);
  const allowedImportKeys = new Set(
    Array.isArray(selectedImportKeys)
      ? selectedImportKeys.map((value) => String(value || "").trim()).filter(Boolean)
      : [],
  );
  if (!allowedImportKeys.size) {
    return {
      summary: {
        file_name: prepared.file_name,
        sheet_name: prepared.sheet_name,
        total_rows: prepared.rows.length,
        selected_rows: 0,
        created: 0,
        updated: 0,
        locked: 0,
        missing_connect: 0,
        skipped: prepared.rows.length,
        failed: 0,
      },
      rows: prepared.rows.map((row) => ({
        ...row,
        imported: false,
        status: row.status,
        note: row.note || "Not selected for import",
      })),
    };
  }
  const selectedRows = prepared.rows.filter((row) => allowedImportKeys.has(String(row.import_key || "").trim()));
  const localActions = await readFustActions();
  const summary = {
    file_name: prepared.file_name,
    sheet_name: prepared.sheet_name,
    total_rows: prepared.rows.length,
    selected_rows: selectedRows.length,
    created: 0,
    updated: 0,
    locked: 0,
    missing_connect: 0,
    skipped: 0,
    failed: 0,
  };
  const results = [];

  for (const row of prepared.rows) {
    if (allowedImportKeys.size && !allowedImportKeys.has(String(row.import_key || "").trim())) {
      summary.skipped += 1;
      results.push({ ...row, imported: false, status: "skipped", note: "Not selected for import" });
      continue;
    }
    if (row.status === "missing_connect") {
      summary.missing_connect += 1;
      results.push({ ...row, imported: false });
      continue;
    }

    try {
      const existingMatch = row.matched_action_id
        ? await ensureLocalFustAction(row.matched_action_id, prepared.settings)
        : { actions: localActions, actionIndex: -1, action: null };
      const existingAction = existingMatch?.action || null;

      if (existingAction && isFustActionConfirmed(existingAction) && !hasUserPermission(requestUser, PERMISSIONS.FUST_MANAGE)) {
        summary.locked += 1;
        results.push({ ...row, imported: false, status: "locked", note: `Confirmed by ${existingAction.confirmed_by || "unknown"}` });
        continue;
      }

      const baseAction = existingAction || {};
      const nextAction = normalizeFustAction({
        ...baseAction,
        id: existingAction?.id || createImportedActionId(row.import_key),
        type: "OUT",
        action_date: row.action_date,
        week: weekNumberForDate(row.action_date),
        day_name: weekdayNameForDate(row.action_date),
        country: row.country,
        customer_name: row.customer_name,
        customer_code: row.customer_code,
        connect_name: row.connect_name,
        remark: baseAction.remark || "",
        fustbon_reference: baseAction.fustbon_reference || "",
        fustfactuur_reference: baseAction.fustfactuur_reference || "",
        metrics: {
          ...(baseAction.metrics || emptyFustMetrics()),
          dc: Number(row.metrics?.dc || 0),
          cctag: Number(row.metrics?.cctag || 0),
          dcs: Number(row.metrics?.dcs || 0),
          dco: Number(row.metrics?.dco || 0),
          pal: Number(row.metrics?.pal || 0),
          vk: Number(row.metrics?.vk || 0),
        },
        created_by: existingAction?.created_by || requestUser.username,
        created_at: existingAction?.created_at || new Date().toISOString(),
        confirmed_at: existingAction?.confirmed_at || "",
        confirmed_by: existingAction?.confirmed_by || "",
        import_source: {
          import_key: row.import_key,
          file_name: prepared.file_name,
          sheet_name: prepared.sheet_name,
          row_number: row.source_row_number,
          source_date: row.action_date,
          imported_at: new Date().toISOString(),
          imported_by: requestUser.username,
          import_kind: "overzicht_out",
        },
        deleted: false,
        deleted_at: "",
        deleted_by: "",
        sheet_sync: { ok: false, target_sheets: [], error: "Pending import sync" },
        email_sync: existingAction?.email_sync?.ok
          ? existingAction.email_sync
          : { ok: true, recipients: [], error: "Imported without email notification" },
        cmr: existingAction?.cmr || normalizeCmrInfo(null),
        fustbon: existingAction?.fustbon || normalizeCmrInfo(null),
      });

      const localIndex = localActions.findIndex((item) => String(item.id || "").trim() === nextAction.id);
      if (localIndex >= 0) {
        localActions[localIndex] = nextAction;
      } else {
        localActions.push(nextAction);
      }
      await writeFustActions(localActions);
      await mirrorFustActionToDatabase(nextAction);

      try {
        nextAction.sheet_sync = await syncFustActionToSheets(nextAction, prepared.settings, { previousAction: existingAction });
      } catch (sheetError) {
        nextAction.sheet_sync = {
          ok: false,
          target_sheets: [],
          error: sheetError instanceof Error ? sheetError.message : String(sheetError),
        };
      }

      const savedIndex = localActions.findIndex((item) => String(item.id || "").trim() === nextAction.id);
      if (savedIndex >= 0) {
        localActions[savedIndex] = nextAction;
      } else {
        localActions.push(nextAction);
      }
      await writeFustActions(localActions);
      await mirrorFustActionToDatabase(nextAction);

      if (existingAction) {
        summary.updated += 1;
      } else {
        summary.created += 1;
      }
      results.push({
        ...row,
        imported: true,
        status: existingAction ? "updated" : "created",
        action_id: nextAction.id,
        sheet_sync: nextAction.sheet_sync,
      });
    } catch (error) {
      summary.failed += 1;
      results.push({
        ...row,
        imported: false,
        status: "failed",
        note: error instanceof Error ? error.message : String(error),
      });
    }
  }

  return {
    summary,
    rows: results,
  };
}

async function backfillFustDatabase() {
  if (!isDatabaseEnabled()) {
    throw new Error("Database is not configured");
  }
  const settings = await readFustSettings();
  const current = await collectCurrentFustActions(settings);
  for (const action of current.activeActions) {
    await saveFustActionToDatabase(action);
  }
  for (const actionId of current.deletedActionIds) {
    await markFustActionDeletedInDatabase(actionId);
  }
  return {
    active_upserted: current.activeActions.length,
    deleted_marked: current.deletedActionIds.length,
    local_action_count: current.localActionCount,
    sheet_action_count: current.sheetActionCount,
    database: await getFustDatabaseStats(),
  };
}

async function findCurrentFustActionById(actionId, settings) {
  const current = await collectCurrentFustActions(settings);
  return current.activeActions.find((action) => String(action.id || "").trim() === String(actionId || "").trim()) || null;
}

async function ensureLocalFustAction(actionId, settings) {
  const normalizedActionId = String(actionId || "").trim();
  const actions = await readFustActions();
  let actionIndex = actions.findIndex((item) => String(item.id || "").trim() === normalizedActionId);
  if (actionIndex >= 0) {
    return {
      actions,
      actionIndex,
      action: actions[actionIndex],
      seededFromSheet: false,
    };
  }

  const currentAction = await findCurrentFustActionById(normalizedActionId, settings);
  if (!currentAction) {
    return {
      actions,
      actionIndex: -1,
      action: null,
      seededFromSheet: false,
    };
  }

  const seededAction = normalizeFustAction({
    ...currentAction,
    created_by: currentAction.created_by === "spreadsheet" ? "sheet-import" : (currentAction.created_by || "sheet-import"),
    created_at: currentAction.created_at || new Date().toISOString(),
  });
  actions.push(seededAction);
  await writeFustActions(actions);
  return {
    actions,
    actionIndex: actions.length - 1,
    action: seededAction,
    seededFromSheet: true,
  };
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
  if (!hasUserPermission(requestUser, permission)) {
    sendJson(res, 403, { error: "You do not have access to this action" });
    return false;
  }
  return true;
}

function requireAnyPermission(res, requestUser, allowedPermissions) {
  const permissions = normalizePermissions(requestUser?.role, requestUser?.permissions);
  if (!Array.isArray(allowedPermissions) || !allowedPermissions.some((permission) => permissions.includes(permission))) {
    sendJson(res, 403, { error: "You do not have access to this action" });
    return false;
  }
  return true;
}

function hasUserPermission(requestUser, permission) {
  return normalizePermissions(requestUser?.role, requestUser?.permissions).includes(permission);
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

function isFustActionConfirmed(action) {
  return Boolean(String(action?.confirmed_at || "").trim());
}

function isImportedFustAction(action) {
  return Boolean(
    String(action?.import_source?.import_key || "").trim()
    || String(action?.import_source?.imported_at || "").trim()
    || String(action?.import_source?.file_name || "").trim(),
  );
}

function buildFustImportKey(actionLike) {
  return [
    String(actionLike?.type || "OUT").trim().toUpperCase(),
    String(actionLike?.action_date || "").trim(),
    String(actionLike?.country || "").trim().toUpperCase(),
    String(actionLike?.customer_name || "").trim().toLowerCase(),
    String(actionLike?.connect_name || actionLike?.customer_code || "").trim().toLowerCase(),
  ].join("|");
}

function createImportedActionId(importKey) {
  const digest = crypto.createHash("sha1").update(String(importKey || "")).digest("hex").slice(0, 16);
  return `fust-import-${digest}`;
}

function matchFustMetaRecord(records, country, customerName) {
  return (records || []).find((record) => (
    String(record.country || "").trim().toUpperCase() === String(country || "").trim().toUpperCase()
    && String(record.customer_name || "").trim().toLowerCase() === String(customerName || "").trim().toLowerCase()
  )) || null;
}

function resolveFustImportMeta(metaRecords, record) {
  const country = String(record?.country || "").trim().toUpperCase();
  const carrier1Name = String(record?.carrier1_name || record?.customer_name || "").trim();
  const carrier2Name = String(record?.carrier2_name || "").trim();
  const carrier1Meta = carrier1Name ? matchFustMetaRecord(metaRecords, country, carrier1Name) : null;
  const carrier2Meta = carrier2Name ? matchFustMetaRecord(metaRecords, country, carrier2Name) : null;

  if (carrier2Meta) {
    return {
      customer_name: carrier2Meta.customer_name,
      connect_name: carrier2Meta.connect_name || carrier2Meta.customer_code || "",
      customer_code: carrier2Meta.customer_code || carrier2Meta.connect_name || "",
      matched_by: "carrier2",
      match_name: carrier2Name,
    };
  }

  if (carrier1Meta) {
    return {
      customer_name: carrier1Meta.customer_name,
      connect_name: carrier1Meta.connect_name || carrier1Meta.customer_code || "",
      customer_code: carrier1Meta.customer_code || carrier1Meta.connect_name || "",
      matched_by: "carrier1",
      match_name: carrier1Name,
    };
  }

  return {
    customer_name: carrier1Name || String(record?.customer_name || "").trim(),
    connect_name: "",
    customer_code: "",
    matched_by: "",
    match_name: "",
  };
}

function findMatchingFustImportAction(actions, candidate) {
  const targetImportKey = buildFustImportKey(candidate);
  const targetBusinessKey = buildActionBusinessKey(candidate);
  return (actions || []).find((action) => {
    if (String(action?.import_source?.import_key || "").trim() && String(action.import_source.import_key).trim() === targetImportKey) {
      return true;
    }
    return buildActionBusinessKey(action) === targetBusinessKey;
  }) || null;
}

async function parseFustImportWorkbook(filePayload) {
  const originalName = path.basename(String(filePayload?.name || "").trim());
  const contentBase64 = String(filePayload?.content_base64 || "").trim();
  if (!originalName || !contentBase64) {
    throw new Error("Choose an import file first");
  }

  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "fust-import-"));
  const extension = path.extname(originalName).toLowerCase() || ".xlsx";
  const inputPath = path.join(tempDir, `source${extension}`);
  try {
    await fs.writeFile(inputPath, Buffer.from(contentBase64, "base64"));
    const output = await runFustImportWorker(["parse", "--input", inputPath]);
    const payload = JSON.parse(output.toString("utf8"));
    return {
      file_name: originalName,
      sheet_name: String(payload?.sheet_name || "Overzicht"),
      records: Array.isArray(payload?.records) ? payload.records : [],
    };
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true }).catch(() => {});
  }
}

async function prepareFustImportRows(filePayload, requestUser) {
  const settings = await readFustSettings();
  const canManageConfirmed = hasUserPermission(requestUser, PERMISSIONS.FUST_MANAGE);
  const [parsed, current, dataRows] = await Promise.all([
    parseFustImportWorkbook(filePayload),
    collectCurrentFustActions(settings),
    loadFustSheetRows(settings),
  ]);
  const meta = buildFustMetaFromSheetRows(dataRows);
  const resolvedRows = parsed.records.map((record) => {
    const resolvedMeta = resolveFustImportMeta(meta.records, record);
    return {
      ...record,
      type: "OUT",
      customer_name: resolvedMeta.customer_name,
      connect_name: resolvedMeta.connect_name,
      customer_code: resolvedMeta.customer_code,
      matched_by: resolvedMeta.matched_by,
      matched_name: resolvedMeta.match_name,
    };
  });

  const groupedRows = new Map();
  for (const record of resolvedRows) {
    const key = [
      "OUT",
      String(record.action_date || "").trim(),
      String(record.country || "").trim().toUpperCase(),
      String(record.customer_name || "").trim().toLowerCase(),
      String(record.connect_name || record.customer_code || "").trim().toLowerCase(),
    ].join("|");
    if (!groupedRows.has(key)) {
      groupedRows.set(key, {
        ...record,
        metrics: { ...record.metrics },
        source_row_numbers: [record.source_row_number].filter(Boolean),
      });
      continue;
    }
    const existing = groupedRows.get(key);
    existing.metrics = {
      dc: Number(existing.metrics?.dc || 0) + Number(record.metrics?.dc || 0),
      cctag: Number(existing.metrics?.cctag || 0) + Number(record.metrics?.cctag || 0),
      dcs: Number(existing.metrics?.dcs || 0) + Number(record.metrics?.dcs || 0),
      dco: Number(existing.metrics?.dco || 0) + Number(record.metrics?.dco || 0),
      pal: Number(existing.metrics?.pal || 0) + Number(record.metrics?.pal || 0),
      vk: Number(existing.metrics?.vk || 0) + Number(record.metrics?.vk || 0),
    };
    existing.source_row_numbers = [...new Set([...(existing.source_row_numbers || []), record.source_row_number].filter(Boolean))];
    if (!existing.matched_by && record.matched_by) {
      existing.matched_by = record.matched_by;
      existing.matched_name = record.matched_name;
    }
  }

  const preparedRows = [...groupedRows.values()].map((record) => {
    const connectName = record.connect_name || "";
    const customerCode = record.customer_code || connectName;
    const importKey = buildFustImportKey({
      type: "OUT",
      action_date: record.action_date,
      country: record.country,
      customer_name: record.customer_name,
      connect_name: connectName,
      customer_code: customerCode,
    });
    const matchedAction = connectName
      ? findMatchingFustImportAction(current.activeActions, {
        type: "OUT",
        action_date: record.action_date,
        country: record.country,
        customer_name: record.customer_name,
        connect_name: connectName,
        customer_code: customerCode,
        import_source: { import_key: importKey },
      })
      : null;
    const confirmed = isFustActionConfirmed(matchedAction);
    const status = !connectName
      ? "missing_connect"
      : matchedAction
        ? (confirmed && !canManageConfirmed ? "locked" : "update")
        : "new";
    const fallbackNote = !connectName && record.carrier2_name
      ? `No connect match for "${record.carrier2_name}". Fallback to "${record.carrier1_name || record.customer_name}" also missing in Fust data sheet`
      : !connectName
        ? "No connect match found in Fust data sheet"
        : "";
    return {
      ...record,
      connect_name: connectName,
      customer_code: customerCode,
      import_key: importKey,
      matched_action_id: matchedAction?.id || "",
      matched_confirmed: confirmed,
      status,
      note: !connectName
        ? fallbackNote
        : matchedAction
          ? (confirmed && !canManageConfirmed ? `Confirmed by ${matchedAction.confirmed_by || "unknown"}` : "Will update existing action")
          : "Will create new action",
    };
  });

  return {
    settings,
    file_name: parsed.file_name,
    sheet_name: parsed.sheet_name,
    rows: preparedRows,
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
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    return raw;
  }
  const match = raw.match(/^(\d{1,2})-(\d{1,2})-(\d{2}|\d{4})$/);
  if (!match) {
    return raw;
  }
  const [, day, month, year] = match;
  const fullYear = year.length === 2 ? `20${year}` : year;
  return `${fullYear}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function sanitizeUkdocsPrintRemark(value) {
  const raw = String(value || "").trim();
  if (!raw) {
    return "";
  }
  if (/^\(?\s*incl\s+ltd\b/i.test(raw)) {
    return "";
  }
  return raw;
}

function parseUkdocsPrintSheetRows(rows, date = localDateIso()) {
  if (!Array.isArray(rows) || rows.length < 2) {
    return [];
  }
  const headers = rows[0].map(normalizeHeader);
  const pick = (aliases, fallbackIndex) => {
    const index = firstMatchingIndex(headers, aliases);
    return index >= 0 ? index : fallbackIndex;
  };
  const dateIndex = pick(["datum", "date"], 1);
  const dayIndex = pick(["dag", "day"], 0);
  const cityIndex = pick(["stad", "city"], 2);
  const borderIndex = pick(["grensovergang", "border crossing"], 3);
  const hubIndex = pick(["hubcode", "hub code", "hub"], 4);
  const remarkIndex = pick(["remark", "opmerking"], 5);
  const pdFormIndex = pick(["pd form", "pd form.", "pd-form"], 7);
  const reExportIndex = pick(["re-export", "re export"], 8);
  const typeIndex = pick(["type"], 9);
  const pdCodeIndex = pick(["code voor pd", "pd code"], 11);
  const referenceConnectIndex = pick(["referentie connect", "reference connect", "connect reference"], 12);
  const trailerIndex = pick(["trailer", "trailer registration", "kenteken trailer"], -1);
  const truckIndex = pick(["truck", "truck registration", "kenteken truck"], -1);
  const invoiceIndex = pick(["invoice", "invoice number", "factuur", "factuurnummer"], -1);

  const parsedRows = rows.slice(1)
    .map((row, index) => {
      const shipmentDate = parseSheetDateToIso(rowValue(row, dateIndex));
      return {
        id: `sheet-row-${shipmentDate}-${rowValue(row, referenceConnectIndex) || index + 2}`,
        shipment_date: shipmentDate,
        day_name: rowValue(row, dayIndex),
        city_name: rowValue(row, cityIndex),
        border_crossing: rowValue(row, borderIndex),
        hub_code: rowValue(row, hubIndex),
        remark: sanitizeUkdocsPrintRemark(rowValue(row, remarkIndex)),
        pd_form: rowValue(row, pdFormIndex),
        re_export: rowValue(row, reExportIndex),
        pd_type: rowValue(row, typeIndex),
        pd_code: rowValue(row, pdCodeIndex),
        reference_connect: rowValue(row, referenceConnectIndex),
        trailer_number: rowValue(row, trailerIndex),
        truck_number: rowValue(row, truckIndex),
        invoice_numbers: rowValue(row, invoiceIndex),
        collection_type: normalizeUkdocsText(rowValue(row, cityIndex)).toUpperCase() === "HONSELERSDIJK" ? "stock_control" : "export",
        sheet_row_number: index + 2,
      };
    })
    .filter((item) => item.shipment_date === String(date || localDateIso()).slice(0, 10))
    .filter((item) => item.reference_connect || item.city_name || item.hub_code);

  const uniqueJoin = (values, separator = "/") => [...new Set(values.map((item) => String(item || "").trim()).filter(Boolean))].join(separator);
  const grouped = new Map();
  for (const item of parsedRows) {
    const groupKey = [
      item.shipment_date,
      normalizeUkdocsPrintToken(item.city_name),
      normalizeUkdocsPrintToken(item.hub_code),
      normalizeUkdocsPrintToken(item.remark),
    ].join("|");
    const existing = grouped.get(groupKey);
    if (!existing) {
      grouped.set(groupKey, {
        ...item,
        id: `sheet-${item.shipment_date}-${sanitizeDriveName(item.city_name || item.hub_code || "shipment")}-${sanitizeDriveName(item.hub_code || "hub")}-${sanitizeDriveName(item.remark || "base")}`,
      });
      continue;
    }
    existing.reference_connect = uniqueJoin([existing.reference_connect, item.reference_connect]);
    existing.invoice_numbers = uniqueJoin([existing.invoice_numbers, item.invoice_numbers]);
    existing.truck_number = uniqueJoin([existing.truck_number, item.truck_number]);
    existing.trailer_number = uniqueJoin([existing.trailer_number, item.trailer_number]);
    existing.pd_form = uniqueJoin([existing.pd_form, item.pd_form], "\n");
    existing.re_export = uniqueJoin([existing.re_export, item.re_export], "\n");
    existing.pd_type = uniqueJoin([existing.pd_type, item.pd_type], "\n");
    existing.pd_code = uniqueJoin([existing.pd_code, item.pd_code], "\n");
    existing.sheet_row_number = Math.min(existing.sheet_row_number || item.sheet_row_number, item.sheet_row_number);
  }
  return [...grouped.values()];
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

function buildActionBusinessKey(action) {
  return [
    String(action?.type || "").trim().toUpperCase(),
    String(action?.action_date || "").trim(),
    String(action?.country || "").trim().toUpperCase(),
    String(action?.customer_name || "").trim().toLowerCase(),
    String(action?.connect_name || action?.customer_code || "").trim().toLowerCase(),
  ].join("|");
}

function buildActionMergeKey(action) {
  const actionId = String(action?.id || "").trim();
  if (actionId) {
    return `id:${actionId}`;
  }
  return `sig:${buildActionSignature(action)}`;
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
      id: rowValue(row, 17) || `sheet-${index + 1}`,
      type: rowValue(row, 0).toLowerCase() === "uitgaand" ? "OUT" : "IN",
      day_name: rowValue(row, 1),
      action_date: parseSheetDateToIso(rowValue(row, 2)),
      week: rowValue(row, 3) ? Number(rowValue(row, 3)) : null,
      customer_name: rowValue(row, 4),
      country: rowValue(row, 5),
      customer_code: rowValue(row, 6),
      connect_name: rowValue(row, 6),
      remark: rowValue(row, 7),
      fustbon_reference: rowValue(row, 15),
      fustfactuur_reference: rowValue(row, 16),
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
      sheet_sync: { ok: true, target_sheets: ["Dashboard"], error: "", row_number: index + (hasHeaderRow ? 2 : 1) },
      email_sync: { ok: true, recipients: [], error: "" },
    }))
    .filter((action) => action.customer_name && action.country);
}

function parseRegistrySheetRows(rows, type) {
  if (!Array.isArray(rows) || rows.length < 2) {
    return [];
  }

  const layout = getRegistrySheetLayout(rows);
  const headers = layout.headers;
  const hasHeaderRow = layout.hasHeaderRow;
  const sourceRows = hasHeaderRow ? rows.slice(1) : rows;

  return sourceRows
    .map((row, index) => normalizeFustAction({
      id: rowValue(row, layout.idIndex) || `${type.toLowerCase()}-sheet-${index + 1}`,
      type,
      day_name: rowValue(row, layout.dayIndex),
      action_date: parseSheetDateToIso(rowValue(row, layout.dateIndex)),
      week: rowValue(row, layout.weekIndex) ? Number(rowValue(row, layout.weekIndex)) : null,
      customer_name: rowValue(row, layout.customerNameIndex),
      country: rowValue(row, layout.countryIndex),
      customer_code: rowValue(row, layout.connectIndex),
      connect_name: rowValue(row, layout.connectIndex),
      remark: rowValue(row, layout.remarkIndex),
      fustbon_reference: rowValue(row, layout.fustbonIndex),
      fustfactuur_reference: rowValue(row, layout.fustfactuurIndex),
      metrics: {
        dc: normalizeNumber(rowValue(row, layout.dcIndex)),
        cctag: normalizeNumber(rowValue(row, layout.cctagIndex)),
        dcs: normalizeNumber(rowValue(row, layout.dcsIndex)),
        dco: normalizeNumber(rowValue(row, layout.dcoIndex)),
        pal: normalizeNumber(rowValue(row, layout.palIndex)),
        vk: normalizeNumber(rowValue(row, layout.vkIndex)),
      },
      created_by: "spreadsheet",
      created_at: "",
      sheet_sync: { ok: true, target_sheets: [type === "OUT" ? "Uitgaand" : "Retour"], error: "", row_number: index + (hasHeaderRow ? 2 : 1) },
      email_sync: { ok: true, recipients: [], error: "" },
    }))
    .filter((action) => action.customer_name && action.country);
}

function getRegistrySheetLayout(rows) {
  const headers = Array.isArray(rows?.[0]) ? rows[0].map(normalizeHeader) : [];
  const hasHeaderRow = headers.includes("klantnaam") || headers.includes("dag");
  const pick = (aliases, fallbackIndex) => {
    const index = firstMatchingIndex(headers, aliases);
    return index >= 0 ? index : fallbackIndex;
  };
  const rowLength = Math.max(
    headers.length,
    pick(["id", "action id", "actie id"], 18) + 1,
    19,
  );
  return {
    headers,
    hasHeaderRow,
    rowLength,
    dayIndex: pick(["dag", "day"], 0),
    dateIndex: pick(["datum", "date"], 1),
    weekIndex: pick(["week"], 2),
    customerNameIndex: pick(["klantnaam", "customer name", "customer"], 3),
    countryIndex: pick(["co", "country", "land"], 4),
    connectIndex: pick(["klantcode connect", "connect", "connect code", "klantcode connector"], 5),
    remarkIndex: pick(["remark", "opmerking"], 6),
    dcIndex: pick(["dc"], 7),
    cctagIndex: pick(["cctag", "cc tag"], 8),
    dcsIndex: pick(["dcs"], 9),
    dcoIndex: pick(["dco"], 10),
    palIndex: pick(["pal", "pallet", "pallets"], 11),
    vkIndex: pick(["vk"], 12),
    fustbonIndex: pick(["fustbon"], 14),
    fustfactuurIndex: pick(["fustfactuur", "fust factuur"], 15),
    idIndex: pick(["id", "action id", "actie id"], 18),
  };
}

function buildRegistrySheetRow(action, rows, existingRow = []) {
  const layout = getRegistrySheetLayout(rows);
  const row = Array.from({ length: layout.rowLength }, (_, index) => String(existingRow?.[index] || ""));
  row[layout.dayIndex] = action.day_name || "";
  row[layout.dateIndex] = isoDateForDisplay(action.action_date);
  row[layout.weekIndex] = action.week ?? "";
  row[layout.customerNameIndex] = action.customer_name || "";
  row[layout.countryIndex] = action.country || "";
  row[layout.connectIndex] = action.customer_code || action.connect_name || "";
  row[layout.remarkIndex] = action.remark || "";
  row[layout.dcIndex] = action.metrics.dc || "";
  row[layout.cctagIndex] = action.metrics.cctag || "";
  row[layout.dcsIndex] = action.metrics.dcs || "";
  row[layout.dcoIndex] = action.metrics.dco || "";
  row[layout.palIndex] = action.metrics.pal || "";
  row[layout.vkIndex] = action.metrics.vk || "";
  row[layout.fustbonIndex] = action.fustbon_reference || "";
  row[layout.fustfactuurIndex] = action.fustfactuur_reference || "";
  row[layout.idIndex] = action.id || "";
  return row;
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
    action.fustbon_reference || "",
    action.fustfactuur_reference || "",
    action.id || "",
  ];
}

function findFustSheetRowNumberByActionId(rows, actionId) {
  if (!actionId || !Array.isArray(rows) || !rows.length) {
    return 0;
  }
  const headers = Array.isArray(rows[0]) ? rows[0].map(normalizeHeader) : [];
  const hasHeaderRow = headers.includes("klantnaam") || headers.includes("dag") || headers.includes("richting");
  const idIndex = firstMatchingIndex(headers, ["id", "action id", "actie id"]);
  const fallbackIdIndex = headers.includes("richting") ? 17 : 16;
  for (const [offset, row] of (hasHeaderRow ? rows.slice(1) : rows).entries()) {
    const candidate = rowValue(row, idIndex >= 0 ? idIndex : fallbackIdIndex);
    if (candidate && candidate === actionId) {
      return offset + (hasHeaderRow ? 2 : 1);
    }
  }
  return 0;
}

function findFustSheetRowNumberBySignature(rows, type, action) {
  if (!Array.isArray(rows) || rows.length < 2 || !action) {
    return 0;
  }
  const parsedRows = parseRegistrySheetRows(rows, type);
  const targetSignature = buildActionSignature(action);
  const matched = parsedRows.find((entry) => buildActionSignature(entry) === targetSignature);
  return Number(matched?.sheet_sync?.row_number || 0);
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

function normalizeSyncStatus(status) {
  const payload = status && typeof status === "object" ? { ...status } : {};
  if (payload.state !== "running") {
    return payload;
  }

  const rawTimestamp = String(payload.updated_at || payload.started_at || "").trim();
  const parsedTimestamp = rawTimestamp ? Date.parse(rawTimestamp) : Number.NaN;
  if (!Number.isFinite(parsedTimestamp)) {
    return payload;
  }

  const ageMs = Date.now() - parsedTimestamp;
  if (ageMs < syncStatusStaleMinutes * 60 * 1000) {
    return payload;
  }

  return {
    ...payload,
    state: "failed",
    stale: true,
    error: String(payload.error || `Sync status became stale after ${syncStatusStaleMinutes} minutes`),
    updated_at: new Date().toISOString(),
  };
}

async function isSyncRunning() {
  const status = normalizeSyncStatus(await readJsonFile(syncStatusPath, {}));
  if (status?.stale) {
    await writeJsonFile(syncStatusPath, status);
  }
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

function runUkdocsWorker(args, input = "") {
  return new Promise((resolve, reject) => {
    const child = spawn(resolvePythonCommand(), [ukdocsWorkerPath, ...args], {
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
      reject(new Error(Buffer.concat(stderr).toString("utf8") || `UKdocs worker exited with ${code}`));
    });

    if (input) {
      child.stdin.write(input);
    }
    child.stdin.end();
  });
}

function runUkdocsCsiWorker(args, input = "") {
  return new Promise((resolve, reject) => {
    const child = spawn(resolvePythonCommand(), [ukdocsCsiWorkerPath, ...args], {
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
      reject(new Error(Buffer.concat(stderr).toString("utf8") || `UKDocs CSI worker exited with ${code}`));
    });

    if (input) {
      child.stdin.write(input);
    }
    child.stdin.end();
  });
}

function mergeUkdocsStatePatch(currentState, patch) {
  return normalizeUkdocsState({
    ...currentState,
    ...patch,
    company_settings: patch.company_settings ? { ...currentState.company_settings, ...patch.company_settings } : currentState.company_settings,
    export_defaults: patch.export_defaults ? { ...currentState.export_defaults, ...patch.export_defaults } : currentState.export_defaults,
  });
}

function ukdocsPrintDocumentPath(document) {
  if (!document?.storage_name) {
    return "";
  }
  return path.join(ukdocsPrintFilesDir, document.storage_name);
}

function upsertUkdocsPrintCollection(collections, nextCollection) {
  const nextShipmentId = normalizeUkdocsText(nextCollection?.shipment_id);
  const existingIndex = collections.findIndex((item) => item.id === nextCollection.id || (nextShipmentId && item.shipment_id === nextShipmentId));
  if (existingIndex >= 0) {
    const nextCollections = [...collections];
    nextCollections[existingIndex] = nextCollection;
    return nextCollections;
  }
  return [nextCollection, ...collections];
}

function ukdocsPrintMatchLines(value) {
  return String(value || "")
    .split(/\r?\n+/)
    .map(normalizeUkdocsPrintToken)
    .filter(Boolean);
}

function matchUkdocsCustomerForPrintCollection(customers, collection) {
  const hubCode = normalizeUkdocsPrintToken(collection?.hub_code);
  const remark = normalizeUkdocsPrintToken(collection?.remark);
  let bestMatch = null;
  let bestScore = 0;
  for (const customer of Array.isArray(customers) ? customers : []) {
    if (!normalizeUkdocsText(customer?.customer_name)) {
      continue;
    }
    const customerHubCodes = ukdocsPrintMatchLines(customer?.match_hub_code);
    const customerRemarks = ukdocsPrintMatchLines(customer?.match_remark);
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

function buildUkdocsPrintCollectionFromShipment(existingCollection, shipment, customerName) {
  return normalizeUkdocsPrintCollection({
    ...existingCollection,
    id: existingCollection?.id || shipment.print_collection_id || shipment.id,
    shipment_id: shipment.id,
    shipment_reference: shipment.export_reference || shipment.invoice_numbers,
    shipment_date: shipment.shipment_date,
    customer_id: shipment.customer_id,
    customer_name: customerName || existingCollection?.customer_name || "",
    collection_type: existingCollection?.collection_type || "export",
    invoice_numbers: shipment.invoice_numbers,
    truck_number: shipment.truck_number,
    trailer_number: shipment.trailer_number,
    reference_connect: shipment.reference_connect || existingCollection?.reference_connect || "",
    generated_at: existingCollection?.generated_at || new Date().toISOString(),
    updated_at: new Date().toISOString(),
    documents: {
      ...(existingCollection?.documents || {}),
      generated_files: normalizeUkdocsPrintDocumentList(existingCollection?.documents?.generated_files),
    },
    notes: existingCollection?.notes || "",
  });
}

async function saveUkdocsPrintUpload(collectionId, kind, filePayload, requestUser) {
  const originalName = path.basename(String(filePayload?.file_name || filePayload?.name || "").trim());
  const contentBase64 = String(filePayload?.content_base64 || "").trim();
  const mimeType = String(filePayload?.mime_type || guessMimeType(originalName)).trim() || "application/octet-stream";
  if (!["phyto", "export_extra", "inspection_list", "locations_file", "temp_phyto", "temp_phyto_plants_file", "temp_phyto_plants_xml_file", "ipaffs_file", "ipaffs_plants_file"].includes(kind)) {
    throw new Error("Unknown UKdocs Print document type");
  }
  if (!originalName || !contentBase64) {
    throw new Error("Choose a file first");
  }

  const extension = safeExtension(originalName, mimeType);
  const fileBuffer = Buffer.from(contentBase64, "base64");
  const storageName = `${sanitizeDriveName(collectionId)}-${kind}-${Date.now()}-${crypto.randomUUID()}${extension}`;
  await fs.mkdir(ukdocsPrintFilesDir, { recursive: true });
  await fs.writeFile(path.join(ukdocsPrintFilesDir, storageName), fileBuffer);
  return {
    storage_name: storageName,
    original_name: originalName,
    mime_type: mimeType,
    size_bytes: fileBuffer.length,
    saved_at: new Date().toISOString(),
    saved_by: requestUser.username,
  };
}

async function saveUkdocsPrintBuffer(collectionId, kind, originalName, mimeType, fileBuffer, savedBy) {
  if (!["phyto", "export_extra", "generated", "inspection_list", "locations_file", "temp_phyto", "temp_phyto_plants_file", "temp_phyto_plants_xml_file", "ipaffs_file", "ipaffs_plants_file"].includes(kind)) {
    throw new Error("Unknown UKdocs Print document type");
  }
  const extension = safeExtension(originalName, mimeType);
  const storageName = `${sanitizeDriveName(collectionId)}-${kind}-${Date.now()}-${crypto.randomUUID()}${extension}`;
  await fs.mkdir(ukdocsPrintFilesDir, { recursive: true });
  await fs.writeFile(path.join(ukdocsPrintFilesDir, storageName), fileBuffer);
  return {
    storage_name: storageName,
    original_name: path.basename(originalName || storageName),
    mime_type: mimeType || guessMimeType(originalName),
    size_bytes: fileBuffer.length,
    saved_at: new Date().toISOString(),
    saved_by: savedBy || "gmail-sync",
  };
}

function hasUkdocsPrintDocumentWithName(documents, originalName) {
  const target = String(originalName || "").trim().toLowerCase();
  if (!target) {
    return false;
  }
  return normalizeUkdocsPrintDocumentList(documents).some((document) => ukdocsPrintDocumentIdentity(document) === target);
}

function buildUkdocsGeneratedFileName(collection, file) {
  const fallbackName = path.basename(String(file?.name || "ukdocs-file").trim()) || "ukdocs-file";
  const documentKind = normalizeUkdocsText(file?.kind);
  if (documentKind !== "export") {
    return fallbackName;
  }

  const extension = path.extname(fallbackName) || ".xlsx";
  const dateText = normalizeUkdocsText(collection?.shipment_date).replace(/-/g, " ").trim();
  const truckText = normalizeUkdocsText(collection?.truck_number);
  const trailerText = normalizeUkdocsText(collection?.trailer_number);
  const referenceText = String(collection?.shipment_reference || collection?.invoice_numbers || "").replace(/[\/\\]+/g, "-").trim();
  const parts = [dateText, truckText, trailerText, referenceText].filter(Boolean);
  if (!parts.length) {
    return fallbackName;
  }
  return `${parts.join(" ")}${extension}`.replace(/\s+/g, " ").trim();
}

async function saveUkdocsGeneratedFiles(collection, files, requestUser) {
  const savedFiles = [];
  for (const file of Array.isArray(files) ? files : []) {
    const contentBase64 = String(file?.content_base64 || "").trim();
    if (!contentBase64) {
      continue;
    }
    const fileBuffer = Buffer.from(contentBase64, "base64");
    const originalName = buildUkdocsGeneratedFileName(collection, file);
    const savedFile = await saveUkdocsPrintBuffer(
      collection?.id,
      "generated",
      originalName,
      String(file?.mime_type || "application/octet-stream"),
      fileBuffer,
      requestUser?.username || "ukdocs-generate",
    );
    savedFiles.push({
      ...savedFile,
      document_kind: normalizeUkdocsText(file?.kind),
      category: normalizeUkdocsText(file?.category),
    });
  }
  return savedFiles;
}

async function queueUkdocsInvoicePdfJobs(collection, requestUser) {
  if (!isDatabaseEnabled() || !llmPollerEnabled()) {
    return [];
  }
  const jobs = [];
  for (const document of Array.isArray(collection?.documents?.generated_files) ? collection.documents.generated_files : []) {
    const mimeType = String(document?.mime_type || "").trim().toLowerCase();
    if (document?.document_kind !== "invoice" || mimeType !== "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") {
      continue;
    }
    const resolvedPath = path.resolve(ukdocsPrintDocumentPath(document));
    if (!resolvedPath.startsWith(path.resolve(ukdocsPrintFilesDir)) || !existsSync(resolvedPath)) {
      continue;
    }
    const workbookBase64 = await fs.readFile(resolvedPath, "base64");
    const pdfName = `${String(document.original_name || "invoice.xlsx").replace(/\.[^.]+$/i, "")}.pdf`;
    const job = await createLlmJob({
      job_type: "excel_to_pdf",
      created_by: requestUser?.username || "ukdocs-generate",
      collection_id: collection.id,
      shipment_id: collection.shipment_id,
      document_kind: "invoice_pdf",
      priority: 40,
      max_attempts: 1,
      payload_json: {
        workbook_content_base64: workbookBase64,
        workbook_name: String(document.original_name || "invoice.xlsx").trim(),
        pdf_name: pdfName,
        source_storage_name: String(document.storage_name || "").trim(),
        category: normalizeUkdocsText(document.category),
      },
    });
    jobs.push(job);
  }
  return jobs;
}

async function saveUkdocsGeneratedInvoicePdfResult(job) {
  const state = await readUkdocsState();
  const existingCollection = ukdocsPrintCollectionById(state.print_collections, job.collection_id);
  if (!existingCollection) {
    return null;
  }

  const exportResult = job?.result_json?.excel_pdf_result || {};
  const sourceStorageName = String(exportResult.source_storage_name || "").trim();
  const category = normalizeUkdocsText(exportResult.category);
  const contentBase64 = String(exportResult.content_base64 || "").trim();
  if (!sourceStorageName || !contentBase64) {
    return null;
  }

  const currentGeneratedFiles = Array.isArray(existingCollection.documents?.generated_files)
    ? [...existingCollection.documents.generated_files]
    : [];
  const sourceStillExists = currentGeneratedFiles.some((file) => String(file?.storage_name || "").trim() === sourceStorageName);
  if (!sourceStillExists) {
    return normalizeUkdocsPrintCollection(existingCollection);
  }

  const filesToRemove = currentGeneratedFiles.filter((file) => (
    file?.document_kind === "invoice"
    && isPdfUkdocsDocument(file)
    && normalizeUkdocsText(file?.category) === category
  ));
  for (const document of filesToRemove) {
    await deleteSingleUkdocsPrintDocumentFile(document);
  }
  const keptFiles = currentGeneratedFiles.filter((file) => !filesToRemove.includes(file));
  const savedFile = await saveUkdocsPrintBuffer(
    existingCollection.id,
    "generated",
    String(exportResult.file_name || `${category || "invoice"}.pdf`).trim(),
    String(exportResult.mime_type || "application/pdf").trim(),
    Buffer.from(contentBase64, "base64"),
    "excel-pdf-poller",
  );
  const updatedCollection = normalizeUkdocsPrintCollection({
    ...existingCollection,
    updated_at: new Date().toISOString(),
    documents: {
      ...(existingCollection.documents || {}),
      generated_files: [
        ...keptFiles,
        {
          ...savedFile,
          document_kind: "invoice",
          category,
        },
      ],
    },
  });
  state.print_collections = upsertUkdocsPrintCollection(state.print_collections, updatedCollection);
  await writeUkdocsState(state);
  return updatedCollection;
}

function normalizeUkdocsPrintToken(value) {
  return String(value || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
}

function ukdocsPrintInvoiceTokens(value) {
  return String(value || "")
    .split(/[\/,\s;]+/)
    .map(normalizeUkdocsPrintToken)
    .filter((item) => item.length >= 4);
}

function countUkdocsGeneratedInvoiceGroups(files) {
  const seen = new Set();
  for (const file of Array.isArray(files) ? files : []) {
    if (file?.document_kind !== "invoice") {
      continue;
    }
    const key = normalizeUkdocsText(file?.category) || String(file?.original_name || file?.storage_name || "").replace(/\.[^.]+$/i, "").trim().toLowerCase();
    if (key) {
      seen.add(key);
    }
  }
  return seen.size;
}

function ukdocsPrintReferenceTokens(value) {
  return String(value || "")
    .split(/[\/,\s;]+/)
    .map(normalizeUkdocsPrintToken)
    .filter((item) => item.length >= 3);
}

function ukdocsPrintCollectionGroupKey(collection) {
  const shipmentDate = normalizeUkdocsText(collection?.shipment_date);
  if (!shipmentDate) {
    return "";
  }
  const collectionType = normalizeUkdocsText(collection?.collection_type) || (isHonselersdijkStockControl(collection) ? "stock_control" : "export");
  return [
    shipmentDate,
    normalizeUkdocsPrintToken(collection?.city_name),
    normalizeUkdocsPrintToken(collection?.hub_code),
    normalizeUkdocsPrintToken(collection?.remark),
    normalizeUkdocsPrintToken(collectionType),
  ].join("|");
}

function mergeUkdocsPrintDocumentLists(primary, secondary) {
  const merged = [];
  const seen = new Set();
  for (const item of [...(Array.isArray(primary) ? primary : []), ...(Array.isArray(secondary) ? secondary : [])]) {
    if (!item) {
      continue;
    }
    const key = `${String(item.storage_name || "").trim()}|${String(item.original_name || "").trim()}`;
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    merged.push(item);
  }
  return merged;
}

function mergeUkdocsPrintCollectionPair(keeper, duplicate) {
  const uniqueJoin = (values, separator = "/") => [...new Set(values.map((item) => String(item || "").trim()).filter(Boolean))].join(separator);
  const chooseObject = (left, right) => (left?.storage_name ? left : (right?.storage_name ? right : null));
  return normalizeUkdocsPrintCollection({
    ...duplicate,
    ...keeper,
    id: keeper.id,
    shipment_id: keeper.shipment_id || duplicate.shipment_id || "",
    source: keeper.source || duplicate.source || "sheet",
    shipment_reference: uniqueJoin([keeper.shipment_reference, duplicate.shipment_reference]),
    customer_id: keeper.customer_id || duplicate.customer_id || "",
    customer_name: keeper.customer_name || duplicate.customer_name || "",
    collection_type: keeper.collection_type || duplicate.collection_type || "export",
    invoice_numbers: uniqueJoin([keeper.invoice_numbers, duplicate.invoice_numbers]),
    truck_number: uniqueJoin([keeper.truck_number, duplicate.truck_number]),
    trailer_number: uniqueJoin([keeper.trailer_number, duplicate.trailer_number]),
    reference_connect: uniqueJoin([keeper.reference_connect, duplicate.reference_connect]),
    city_name: keeper.city_name || duplicate.city_name || "",
    border_crossing: keeper.border_crossing || duplicate.border_crossing || "",
    hub_code: keeper.hub_code || duplicate.hub_code || "",
    remark: keeper.remark || duplicate.remark || "",
    pd_form: uniqueJoin([keeper.pd_form, duplicate.pd_form], "\n"),
    re_export: uniqueJoin([keeper.re_export, duplicate.re_export], "\n"),
    pd_type: uniqueJoin([keeper.pd_type, duplicate.pd_type], "\n"),
    pd_code: uniqueJoin([keeper.pd_code, duplicate.pd_code], "\n"),
    sheet_row_number: Math.min(
      Number(keeper.sheet_row_number || 0) || Number.MAX_SAFE_INTEGER,
      Number(duplicate.sheet_row_number || 0) || Number.MAX_SAFE_INTEGER,
    ) === Number.MAX_SAFE_INTEGER ? 0 : Math.min(
      Number(keeper.sheet_row_number || 0) || Number.MAX_SAFE_INTEGER,
      Number(duplicate.sheet_row_number || 0) || Number.MAX_SAFE_INTEGER,
    ),
    generated_at: keeper.generated_at || duplicate.generated_at || "",
    updated_at: new Date().toISOString(),
    notes: keeper.notes || duplicate.notes || "",
    delivery_email: keeper.delivery_email?.ok ? keeper.delivery_email : (duplicate.delivery_email || keeper.delivery_email || {}),
    csi_email: keeper.csi_email?.ok ? keeper.csi_email : (duplicate.csi_email || keeper.csi_email || {}),
    csi_report: keeper.csi_report?.status ? keeper.csi_report : (duplicate.csi_report || keeper.csi_report || {}),
    documents: {
      phyto_files: mergeUkdocsPrintDocumentLists(keeper?.documents?.phyto_files, duplicate?.documents?.phyto_files),
      export_extra: chooseObject(keeper?.documents?.export_extra, duplicate?.documents?.export_extra),
      generated_files: mergeUkdocsPrintDocumentLists(keeper?.documents?.generated_files, duplicate?.documents?.generated_files),
      inspection_list: chooseObject(keeper?.documents?.inspection_list, duplicate?.documents?.inspection_list),
      locations_file: chooseObject(keeper?.documents?.locations_file, duplicate?.documents?.locations_file),
      temp_phyto_files: mergeUkdocsPrintDocumentLists(keeper?.documents?.temp_phyto_files, duplicate?.documents?.temp_phyto_files),
      temp_phyto_plants_file: chooseObject(keeper?.documents?.temp_phyto_plants_file, duplicate?.documents?.temp_phyto_plants_file),
      temp_phyto_plants_xml_file: chooseObject(keeper?.documents?.temp_phyto_plants_xml_file, duplicate?.documents?.temp_phyto_plants_xml_file),
      ipaffs_file: chooseObject(keeper?.documents?.ipaffs_file, duplicate?.documents?.ipaffs_file),
      ipaffs_plants_file: chooseObject(keeper?.documents?.ipaffs_plants_file, duplicate?.documents?.ipaffs_plants_file),
    },
  });
}

function ukdocsPrintCollectionMergeScore(collection) {
  const documents = collection?.documents || {};
  return (
    (collection?.shipment_id ? 50 : 0)
    + ((documents.generated_files || []).length * 10)
    + ((documents.phyto_files || []).length * 5)
    + ((documents.temp_phyto_files || []).length * 5)
    + (documents.temp_phyto_plants_file?.storage_name ? 5 : 0)
    + (documents.temp_phyto_plants_xml_file?.storage_name ? 3 : 0)
    + (documents.export_extra?.storage_name ? 4 : 0)
    + (documents.inspection_list?.storage_name ? 4 : 0)
    + (documents.locations_file?.storage_name ? 4 : 0)
    + (documents.ipaffs_file?.storage_name ? 4 : 0)
    + (documents.ipaffs_plants_file?.storage_name ? 4 : 0)
    + (collection?.delivery_email?.ok ? 8 : 0)
    + (collection?.csi_email?.ok ? 8 : 0)
    + (collection?.csi_report?.status === "done" ? 8 : 0)
    + (ukdocsPrintReferenceTokens(collection?.reference_connect).length * 3)
    + (ukdocsPrintInvoiceTokens(collection?.invoice_numbers).length * 2)
    + (normalizeUkdocsPrintToken(collection?.truck_number) ? 2 : 0)
    + (normalizeUkdocsPrintToken(collection?.trailer_number) ? 2 : 0)
    + (normalizeUkdocsPrintToken(collection?.customer_name) ? 1 : 0)
  );
}

function dedupeUkdocsPrintCollectionsForDate(state, syncDate) {
  const collections = Array.isArray(state?.print_collections) ? state.print_collections : [];
  const shipments = Array.isArray(state?.shipments) ? state.shipments : [];
  const groups = new Map();
  const addBucket = (key, collection) => {
    if (!key) {
      return;
    }
    const bucket = groups.get(key) || [];
    bucket.push(collection);
    groups.set(key, bucket);
  };
  for (const collection of collections) {
    if (String(collection?.shipment_date || "").slice(0, 10) !== syncDate) {
      continue;
    }
    if (collection?.collection_type === "stock_control") {
      continue;
    }
    const rowNumber = Number(collection?.sheet_row_number || 0);
    if (rowNumber > 0) {
      addBucket(`row|${syncDate}|${rowNumber}`, collection);
    }
    addBucket(`group|${ukdocsPrintCollectionGroupKey(collection)}`, collection);
  }

  if (![...groups.values()].some((bucket) => bucket.length > 1)) {
    return state;
  }

  const idRemap = new Map();
  const replacementById = new Map();

  for (const bucket of groups.values()) {
    if (bucket.length < 2) {
      continue;
    }
    const sorted = [...bucket].sort((left, right) => {
      const scoreDiff = ukdocsPrintCollectionMergeScore(right) - ukdocsPrintCollectionMergeScore(left);
      if (scoreDiff !== 0) {
        return scoreDiff;
      }
      return String(left.id || "").localeCompare(String(right.id || ""));
    });
    let keeper = sorted[0];
    for (const duplicate of sorted.slice(1)) {
      keeper = mergeUkdocsPrintCollectionPair(keeper, duplicate);
      idRemap.set(duplicate.id, keeper.id);
      replacementById.set(duplicate.id, keeper);
    }
    replacementById.set(keeper.id, keeper);
  }

  if (!idRemap.size) {
    return state;
  }

  state.print_collections = collections
    .filter((item) => !idRemap.has(item.id))
    .map((item) => replacementById.get(item.id) || item);

  state.shipments = shipments.map((shipment) => {
    const remappedCollectionId = idRemap.get(shipment?.print_collection_id);
    if (!remappedCollectionId) {
      return shipment;
    }
    return normalizeUkdocsShipment({
      ...shipment,
      print_collection_id: remappedCollectionId,
      updated_at: new Date().toISOString(),
    });
  });

  return state;
}

function findMatchingUkdocsPrintCollection(collections, candidate, options = {}) {
  const allowInvoiceFallback = options.allowInvoiceFallback !== false;
  const shipmentId = normalizeUkdocsText(candidate?.shipment_id);
  const collectionId = normalizeUkdocsText(candidate?.id);
  const shipmentDate = normalizeUkdocsText(candidate?.shipment_date);
  const sheetRowNumber = Number(candidate?.sheet_row_number || 0);
  const referenceTokens = ukdocsPrintReferenceTokens(candidate?.reference_connect);
  const invoiceTokens = ukdocsPrintInvoiceTokens(candidate?.invoice_numbers);
  const truckNumber = normalizeUkdocsPrintToken(candidate?.truck_number);
  const trailerNumber = normalizeUkdocsPrintToken(candidate?.trailer_number);
  const candidateGroupKey = ukdocsPrintCollectionGroupKey(candidate);

  return (Array.isArray(collections) ? collections : []).find((item) => {
    if (collectionId && item.id === collectionId) {
      return true;
    }
    if (shipmentId && item.shipment_id === shipmentId) {
      return true;
    }
    if (!shipmentDate || item.shipment_date !== shipmentDate) {
      return false;
    }

    const itemSheetRowNumber = Number(item?.sheet_row_number || 0);
    if (sheetRowNumber > 0 && itemSheetRowNumber > 0 && sheetRowNumber === itemSheetRowNumber) {
      return true;
    }

    const itemReferenceTokens = ukdocsPrintReferenceTokens(item.reference_connect);
    if (referenceTokens.length && itemReferenceTokens.some((token) => referenceTokens.includes(token))) {
      return true;
    }

    const itemGroupKey = ukdocsPrintCollectionGroupKey(item);
    if (candidateGroupKey && itemGroupKey && candidateGroupKey === itemGroupKey) {
      return true;
    }

    if (!allowInvoiceFallback) {
      return false;
    }

    if (!invoiceTokens.length) {
      return false;
    }

    const itemInvoiceTokens = ukdocsPrintInvoiceTokens(item.invoice_numbers);
    if (!itemInvoiceTokens.length) {
      return false;
    }

    const invoiceOverlap = invoiceTokens.some((token) => itemInvoiceTokens.includes(token));
    if (!invoiceOverlap) {
      return false;
    }

    const itemTruckNumber = normalizeUkdocsPrintToken(item.truck_number);
    const itemTrailerNumber = normalizeUkdocsPrintToken(item.trailer_number);
    if (truckNumber && itemTruckNumber && truckNumber === itemTruckNumber) {
      return true;
    }
    if (trailerNumber && itemTrailerNumber && trailerNumber === itemTrailerNumber) {
      return true;
    }

    return !referenceTokens.length && !itemReferenceTokens.length;
  }) || null;
}

async function deleteUkdocsPrintCollectionFiles(collection) {
  const documents = [];
  if (Array.isArray(collection?.documents?.phyto_files)) {
    documents.push(...collection.documents.phyto_files);
  }
  if (collection?.documents?.export_extra) {
    documents.push(collection.documents.export_extra);
  }
  if (Array.isArray(collection?.documents?.generated_files)) {
    documents.push(...collection.documents.generated_files);
  }
  if (collection?.documents?.inspection_list) {
    documents.push(collection.documents.inspection_list);
  }
  if (collection?.documents?.locations_file) {
    documents.push(collection.documents.locations_file);
  }
  if (Array.isArray(collection?.documents?.temp_phyto_files)) {
    documents.push(...collection.documents.temp_phyto_files);
  }
  if (collection?.documents?.temp_phyto_plants_file) {
    documents.push(collection.documents.temp_phyto_plants_file);
  }
  if (collection?.documents?.temp_phyto_plants_xml_file) {
    documents.push(collection.documents.temp_phyto_plants_xml_file);
  }
  if (collection?.documents?.ipaffs_file) {
    documents.push(collection.documents.ipaffs_file);
  }
  if (collection?.documents?.ipaffs_plants_file) {
    documents.push(collection.documents.ipaffs_plants_file);
  }
  await Promise.all(documents.map(async (document) => {
    const resolvedPath = path.resolve(ukdocsPrintDocumentPath(document));
    if (!resolvedPath.startsWith(path.resolve(ukdocsPrintFilesDir))) {
      return;
    }
    try {
      await fs.unlink(resolvedPath);
    } catch (error) {
      if (error?.code !== "ENOENT") {
        throw error;
      }
    }
  }));
}

async function deleteSingleUkdocsPrintDocumentFile(document) {
  if (!document?.storage_name) {
    return;
  }
  const resolvedPath = path.resolve(ukdocsPrintDocumentPath(document));
  if (!resolvedPath.startsWith(path.resolve(ukdocsPrintFilesDir))) {
    return;
  }
  try {
    await fs.unlink(resolvedPath);
  } catch (error) {
    if (error?.code !== "ENOENT") {
      throw error;
    }
  }
}

function ukdocsPrintCollectionById(collections, collectionId) {
  return (collections || []).find((item) => item.id === collectionId || item.shipment_id === collectionId) || null;
}

function extractJsonObjectFromText(value) {
  const text = String(value || "").trim();
  if (!text) {
    return null;
  }
  try {
    return JSON.parse(text);
  } catch {
    const fenced = text.match(/```(?:json)?\s*([\s\S]+?)\s*```/i);
    if (fenced?.[1]) {
      try {
        return JSON.parse(fenced[1]);
      } catch {
        return null;
      }
    }
    const firstBrace = text.indexOf("{");
    const lastBrace = text.lastIndexOf("}");
    if (firstBrace >= 0 && lastBrace > firstBrace) {
      try {
        return JSON.parse(text.slice(firstBrace, lastBrace + 1));
      } catch {
        return null;
      }
    }
    return null;
  }
}

function normalizeUkdocsCsiParsedResult(source) {
  if (!source || typeof source !== "object") {
    return null;
  }
  const checks = Array.isArray(source.checks) ? source.checks : [];
  const products = Array.isArray(source.products) ? source.products : [];
  const flowerProducts = Array.isArray(source.flower_products) ? source.flower_products : [];
  const plantProducts = Array.isArray(source.plant_products) ? source.plant_products : [];
  const sourceRows = Array.isArray(source.source_rows) ? source.source_rows : [];
  const visibleDocuments = Array.isArray(source.visible_documents)
    ? source.visible_documents
    : (Array.isArray(source.visibleDocuments) ? source.visibleDocuments : []);
  const manualChecks = Array.isArray(source.manual_checks)
    ? source.manual_checks
    : (Array.isArray(source.manualChecks) ? source.manualChecks : []);
  const notes = Array.isArray(source.notes) ? source.notes : [];
  const summary = String(source.summary || "").trim();
  const overallStatus = normalizeUkdocsText(source.overall_status || source.overallStatus);
  if (!summary && !checks.length && !products.length && !flowerProducts.length && !plantProducts.length && !sourceRows.length && !visibleDocuments.length && !manualChecks.length && !notes.length && !overallStatus) {
    return null;
  }
  return {
    overall_status: overallStatus || "warn",
    summary: summary || "CSI audit completed.",
    checks,
    products,
    flower_products: flowerProducts,
    plant_products: plantProducts,
    source_rows: sourceRows,
    visible_documents: visibleDocuments,
    manual_checks: manualChecks,
    notes,
  };
}

async function extractUkdocsCsiFileSnapshots(files) {
  const safeFiles = [];
  const skippedFiles = [];

  for (const file of Array.isArray(files) ? files : []) {
    if (!file?.document?.storage_name) {
      continue;
    }
    const resolvedPath = path.resolve(ukdocsPrintDocumentPath(file.document));
    const kind = String(file.kind || "").trim();
    const name = String(file.document.original_name || file.document.storage_name || "").trim();
    const mimeType = String(file.document.mime_type || "").trim();
    if (!resolvedPath.startsWith(path.resolve(ukdocsPrintFilesDir))) {
      skippedFiles.push({
        kind,
        name,
        mime_type: mimeType,
        content_type: "missing",
        text: "",
        line_count: 0,
        error: "Resolved CSI document path is outside the UKDocs storage folder.",
      });
      continue;
    }
    if (!existsSync(resolvedPath)) {
      skippedFiles.push({
        kind,
        name,
        mime_type: mimeType,
        content_type: "missing",
        text: "",
        line_count: 0,
        error: "CSI document file is missing on disk.",
      });
      continue;
    }
    safeFiles.push({
      kind,
      name,
      mime_type: mimeType,
      path: resolvedPath,
    });
  }

  if (!safeFiles.length) {
    return skippedFiles;
  }

  const output = await runUkdocsCsiWorker(["extract"], JSON.stringify({ files: safeFiles }));
  const payload = JSON.parse(output.toString("utf8"));
  return [
    ...(Array.isArray(payload?.documents) ? payload.documents : []),
    ...skippedFiles,
  ];
}

async function enrichUkdocsCsiStoredDocument(document, kind) {
  const normalizedDocument = normalizeUkdocsPrintDocument(document);
  if (!normalizedDocument?.storage_name || !kind) {
    return normalizedDocument;
  }
  const extracted = await extractUkdocsCsiFileSnapshots([{ kind, document: normalizedDocument }]);
  const snapshot = (Array.isArray(extracted) ? extracted : []).find((item) => String(item?.kind || "").trim() === String(kind).trim()) || null;
  if (!snapshot) {
    return normalizedDocument;
  }
  return normalizeUkdocsPrintDocument({
    ...normalizedDocument,
    content_type: snapshot.content_type || normalizedDocument.content_type || "",
    line_count: Number.isFinite(Number(snapshot.line_count)) ? Number(snapshot.line_count) : 0,
    delimiter: String(snapshot.delimiter || snapshot?.parsed_data?.delimiter || "").trim(),
    parse_error: String(snapshot.error || "").trim(),
    parsed_data: snapshot.parsed_data && typeof snapshot.parsed_data === "object" ? snapshot.parsed_data : null,
  });
}

async function hydrateUkdocsCsiCollectionInputs(collection, options = {}) {
  const normalizedCollection = normalizeUkdocsPrintCollection(collection);
  if (!normalizedCollection) {
    return { collection: normalizedCollection, changed: false };
  }
  const forceRefresh = options?.force_refresh === true;
  let changed = false;
  const nextDocuments = { ...(normalizedCollection.documents || {}) };

  if (nextDocuments.ipaffs_file?.storage_name && (forceRefresh || !nextDocuments.ipaffs_file?.parsed_data)) {
    nextDocuments.ipaffs_file = await enrichUkdocsCsiStoredDocument(nextDocuments.ipaffs_file, "ipaffs_file");
    changed = true;
  }
  if (nextDocuments.ipaffs_plants_file?.storage_name && (forceRefresh || !nextDocuments.ipaffs_plants_file?.parsed_data)) {
    nextDocuments.ipaffs_plants_file = await enrichUkdocsCsiStoredDocument(nextDocuments.ipaffs_plants_file, "ipaffs_plants_file");
    changed = true;
  }

  const tempPhytoFiles = [];
  for (const document of nextDocuments.temp_phyto_files || []) {
    if (document?.storage_name && (forceRefresh || !document?.parsed_data)) {
      tempPhytoFiles.push(await enrichUkdocsCsiStoredDocument(document, "temp_phyto"));
      changed = true;
    } else {
      tempPhytoFiles.push(document);
    }
  }
  nextDocuments.temp_phyto_files = tempPhytoFiles;
  if (nextDocuments.temp_phyto_plants_file?.storage_name && (forceRefresh || !nextDocuments.temp_phyto_plants_file?.parsed_data)) {
    nextDocuments.temp_phyto_plants_file = await enrichUkdocsCsiStoredDocument(nextDocuments.temp_phyto_plants_file, "temp_phyto_plants_file");
    changed = true;
  }
  if (nextDocuments.temp_phyto_plants_xml_file?.storage_name && (forceRefresh || !nextDocuments.temp_phyto_plants_xml_file?.parsed_data)) {
    nextDocuments.temp_phyto_plants_xml_file = await enrichUkdocsCsiStoredDocument(nextDocuments.temp_phyto_plants_xml_file, "temp_phyto_plants_xml_file");
    changed = true;
  }

  if (!changed) {
    return { collection: normalizedCollection, changed: false };
  }

  return {
    collection: normalizeUkdocsPrintCollection({
      ...normalizedCollection,
      updated_at: new Date().toISOString(),
      documents: nextDocuments,
    }),
    changed: true,
  };
}

function shortenUkdocsCsiText(value, maxLength = 1200) {
  const text = String(value || "").trim();
  if (!text) {
    return "";
  }
  if (text.length <= maxLength) {
    return text;
  }
  return `${text.slice(0, maxLength)}\n... truncated ${text.length - maxLength} more characters`;
}

function summarizeUkdocsCsiExtractedDocuments(extractedDocuments) {
  return (Array.isArray(extractedDocuments) ? extractedDocuments : []).map((document) => {
    const parsedData = document?.parsed_data && typeof document.parsed_data === "object"
      ? document.parsed_data
      : null;
    const summary = {
      kind: String(document?.kind || "").trim(),
      name: String(document?.name || "").trim(),
      content_type: String(document?.content_type || "").trim(),
      line_count: Number.isFinite(Number(document?.line_count)) ? Number(document.line_count) : 0,
      parsed_data: parsedData,
      problems: Array.isArray(document?.problems) ? document.problems : [],
      error: String(document?.error || "").trim(),
    };

    const kind = summary.kind;
    if (!parsedData) {
      summary.text_excerpt = shortenUkdocsCsiText(document?.text || "", 1800);
      return summary;
    }

    if (kind === "generated_invoice") {
      const invoiceRows = Array.isArray(parsedData?.rows) ? parsedData.rows.slice(0, 80) : [];
      const productOriginTotals = Array.isArray(parsedData?.product_origin_totals) ? parsedData.product_origin_totals.slice(0, 80) : [];
      summary.parsed_data = {
        meta: parsedData?.meta && typeof parsedData.meta === "object" ? parsedData.meta : {},
        rows: invoiceRows,
        product_origin_totals: productOriginTotals,
      };
      return summary;
    }

    if (kind === "generated_export") {
      const exportRows = Array.isArray(parsedData?.rows) ? parsedData.rows.slice(0, 120) : [];
      const productOriginTotals = Array.isArray(parsedData?.product_origin_totals) ? parsedData.product_origin_totals.slice(0, 120) : [];
      summary.parsed_data = {
        rows: exportRows,
        product_origin_totals: productOriginTotals,
      };
      return summary;
    }

    if (kind === "ipaffs_file" || kind === "ipaffs_plants_file") {
      const ipaffsRows = Array.isArray(parsedData?.rows) ? parsedData.rows.slice(0, 120) : [];
      const productTotals = Array.isArray(parsedData?.product_totals) ? parsedData.product_totals.slice(0, 120) : [];
      summary.parsed_data = {
        rows: ipaffsRows,
        product_totals: productTotals,
      };
      summary.ipaffs_debug = {
        delimiter: parsedData?.delimiter || document?.delimiter || "",
        row_count: Array.isArray(parsedData?.rows) ? parsedData.rows.length : 0,
        line_count: summary.line_count,
      };
      return summary;
    }

    if (kind === "temp_phyto" || kind === "temp_phyto_plants_file" || kind === "temp_phyto_plants_xml_file") {
      const productLines = Array.isArray(parsedData?.product_lines) ? parsedData.product_lines.slice(0, 60) : [];
      summary.parsed_data = {
        document_state: parsedData?.document_state || "",
        pcnu_number: parsedData?.pcnu_number || "",
        destination_country: parsedData?.destination_country || "",
        origin_country: parsedData?.origin_country || "",
        consignee: parsedData?.consignee || "",
        total_quantity: parsedData?.total_quantity ?? null,
        product_lines: productLines,
        problems: Array.isArray(parsedData?.problems) ? parsedData.problems : [],
      };
      return summary;
    }

    summary.text_excerpt = shortenUkdocsCsiText(document?.text || "", 1000);
    return summary;
  });
}

function normalizeUkdocsCsiToken(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function normalizeUkdocsCsiOrigin(value) {
  return String(value || "").trim().toUpperCase();
}

const UKDOCS_CSI_PLANT_GROUPS = new Set([
  "CITES ge. non-flowering p",
  "Other non-flowering plant",
  "Flowering plants(no cactu",
  "Perennials",
  "Others",
  "refined roses",
]);

const UKDOCS_CSI_FLOWER_GROUPS = new Set([
  "Flowers (other fresh)",
  "Flowers carnation",
  "Flowers chrysanthemums",
  "Flowers green",
  "Flowers lilies",
  "Flowers orchids",
  "Flowers roses",
]);

function getUkdocsCsiProductDomain(productName) {
  const product = String(productName || "").trim();
  if (UKDOCS_CSI_PLANT_GROUPS.has(product)) {
    return "plants";
  }
  if (UKDOCS_CSI_FLOWER_GROUPS.has(product)) {
    return "flowers";
  }
  return "";
}

function inferUkdocsCsiDocumentDomain(documentName = "", fallbackKind = "") {
  const nameToken = normalizeUkdocsCsiToken(documentName);
  const kindToken = normalizeUkdocsCsiToken(fallbackKind);
  if (
    nameToken.includes("planten")
    || nameToken.includes("plants")
    || kindToken.includes("plants")
  ) {
    return "plants";
  }
  if (
    nameToken.includes("bloemen")
    || nameToken.includes("flowers")
    || kindToken.includes("flower")
  ) {
    return "flowers";
  }
  return "";
}

function normalizeUkdocsCsiKnownGroup(value) {
  const text = normalizeUkdocsCsiToken(value);
  if (!text) {
    return "";
  }
  if (text.includes("cites ge non flowering")) {
    return "CITES ge. non-flowering p";
  }
  if (text.includes("other non flowering plant")) {
    return "Other non-flowering plant";
  }
  if (text.includes("cites flowering plants") || text.includes("flowering plants no cactu")) {
    return "Flowering plants(no cactu";
  }
  if (text.includes("perennials")) {
    return "Perennials";
  }
  if (text === "others" || text.endsWith(" others")) {
    return "Others";
  }
  if (text.includes("refined roses")) {
    return "refined roses";
  }
  if (text.includes("flowers other fresh")) {
    return "Flowers (other fresh)";
  }
  if (text.includes("flowers carnation") || text.includes("flowers carnations")) {
    return "Flowers carnation";
  }
  if (text.includes("flowers chrysanthem")) {
    return "Flowers chrysanthemums";
  }
  if (text.includes("flowers roses")) {
    return "Flowers roses";
  }
  if (text.includes("flowers lilies")) {
    return "Flowers lilies";
  }
  if (text.includes("flowers orchids")) {
    return "Flowers orchids";
  }
  if (text.includes("flowers green")) {
    return "Flowers green";
  }
  return "";
}

function hasAnyUkdocsCsiToken(text, tokens) {
  return tokens.some((token) => text.includes(token));
}

function mapUkdocsCsiAmbiguousPlantGroup(text) {
  if (hasAnyUkdocsCsiToken(text, ["aloe", "curio", "crassula", "echeveria", "succulent", "cactus", "rhipsalis", "sageretia", "bonsai"])) {
    return "CITES ge. non-flowering p";
  }
  if (hasAnyUkdocsCsiToken(text, ["chlorophytum", "dracaena", "dypsis", "epipremnum", "fittonia", "maranta", "nephrolepis", "schefflera", "spathiphyllum", "zamioculcas", "sansevieria", "sanseveria"])) {
    return "Other non-flowering plant";
  }
  return "";
}

function inferUkdocsCsiPlantsPreference(document, fallbackValue = false) {
  if (document?.prefer_plants === true) {
    return true;
  }
  const documentName = normalizeUkdocsCsiToken(document?.name || document?.original_name || "");
  if (documentName.includes("plants")) {
    return true;
  }
  const productLines = Array.isArray(document?.parsed_data?.product_lines) ? document.parsed_data.product_lines : [];
  if (!productLines.length) {
    return fallbackValue;
  }
  let plantScore = 0;
  let flowerScore = 0;
  for (const line of productLines) {
    const text = normalizeUkdocsCsiToken(line?.product || "");
    if (!text) {
      continue;
    }
    if (
      text.includes("aloe") || text.includes("anthurium") || text.includes("begonia")
      || text.includes("bonsai") || text.includes("calathea") || text.includes("campanula")
      || text.includes("celosia") || text.includes("chlorophytum") || text.includes("crassula")
      || text.includes("curio") || text.includes("cyclamen") || text.includes("dracaena")
      || text.includes("dypsis") || text.includes("echeveria") || text.includes("epipremnum")
      || text.includes("fittonia") || text.includes("fuchsia") || text.includes("gerbera")
      || text.includes("guzmania") || text.includes("helianthus") || text.includes("helleanthus")
      || text.includes("hibiscus") || text.includes("hydrangea") || text.includes("kalanchoe")
      || text.includes("lavandula") || text.includes("lithodora") || text.includes("mandevilla")
      || text.includes("maranta") || text.includes("nephrolepis") || text.includes("phalaenopsis")
      || text.includes("platycodon") || text.includes("rhipsalis") || text.includes("sageretia")
      || text.includes("schefflera") || text.includes("spathiphyllum") || text.includes("succulent")
      || text.includes("cactus") || text.includes("sansevieria") || text.includes("sanseveria")
      || text.includes("zamioculcas")
    ) {
      plantScore += 2;
    }
    if (
      text.includes("solidago") || text.includes("gypsoph") || text.includes("carnation")
      || text.includes("other fresh") || text.includes("cutflowers") || text.includes("cut flowers")
      || text.includes("branch")
    ) {
      flowerScore += 2;
    }
    if (text.includes("chrysanthem") || text.includes("dianthus") || text.includes("rosa") || text.includes("rose")) {
      plantScore += 1;
      flowerScore += 1;
    }
  }
  return plantScore > flowerScore || fallbackValue;
}

function mapUkdocsCsiProductName(description, commodityCode = "", options = {}) {
  const text = normalizeUkdocsCsiToken(description);
  const code = String(commodityCode || "").replace(/\D+/g, "");
  const strictDomain = options?.strict_domain === "plants"
    ? "plants"
    : options?.strict_domain === "flowers"
      ? "flowers"
      : "";
  const preferPlants = strictDomain
    ? strictDomain === "plants"
    : (
      options?.prefer_plants === true
      || normalizeUkdocsCsiToken(options?.document_name || "").includes("plants")
    );
  const knownGroup = normalizeUkdocsCsiKnownGroup(description);

  if (!strictDomain && knownGroup) {
    return knownGroup;
  }

  if (strictDomain === "flowers") {
    if (knownGroup && UKDOCS_CSI_FLOWER_GROUPS.has(knownGroup)) {
      return knownGroup;
    }
    if (text.includes("chrysanthem") || code.startsWith("060314") || code.startsWith("603140")) {
      return "Flowers chrysanthemums";
    }
    if (text.includes("dianthus") || text.includes("carnation") || code.startsWith("060312") || code.startsWith("603120")) {
      return "Flowers carnation";
    }
    if (text.includes("rosa") || text.includes("rose") || code.startsWith("060311") || code.startsWith("603110")) {
      return "Flowers roses";
    }
    if (text.includes("lil") || text.includes("lilium") || code.startsWith("060315") || code.startsWith("603150")) {
      return "Flowers lilies";
    }
    if (text.includes("orchid") || text.includes("dendrob") || code.startsWith("060313") || code.startsWith("603130")) {
      return "Flowers orchids";
    }
    if (text.includes("green") || code.startsWith("0604209") || code.startsWith("604209")) {
      return "Flowers green";
    }
    return "Flowers (other fresh)";
  }

  if (strictDomain === "plants") {
    if (code.startsWith("060240") || code.startsWith("60240")) {
      return "refined roses";
    }
    if (code.startsWith("060290500") || code.startsWith("60290500") || code.startsWith("6029050")) {
      return "Perennials";
    }
    if (code.startsWith("060319700") || code.startsWith("60319700") || code === "6031970") {
      return "Others";
    }
    if (code.startsWith("06029091") || code.startsWith("6029091")) {
      return "Flowering plants(no cactu";
    }
    if (knownGroup && UKDOCS_CSI_PLANT_GROUPS.has(knownGroup)) {
      return knownGroup;
    }
    if (
      code.startsWith("060290990") || code.startsWith("60290990")
      || code.startsWith("060290991") || code.startsWith("60290991")
    ) {
      const ambiguousGroup = mapUkdocsCsiAmbiguousPlantGroup(text);
      if (ambiguousGroup) {
        return ambiguousGroup;
      }
      if (knownGroup === "CITES ge. non-flowering p" || knownGroup === "Other non-flowering plant") {
        return knownGroup;
      }
    }
    if (text.includes("bonsai") || text.includes("sageretia") || text.includes("aloe") || text.includes("rhipsalis")) {
      return "CITES ge. non-flowering p";
    }
    if (text.includes("salvia") || text.includes("lavandula") || text.includes("helleborus") || text.includes("campanula")) {
      return "Perennials";
    }
    if (text.includes("hibiscus")) {
      return "Flowering plants(no cactu";
    }
    if (text.includes("cupressus")) {
      return "Perennials";
    }
    if (text.includes("ficus")) {
      return "Others";
    }
    if (text.includes("rosa") || (text.includes("rose") && !text.includes("hibiscus"))) {
      return "refined roses";
    }
    if (
      text.includes("dypsis") || text.includes("maranta") || text.includes("calathea")
      || text.includes("chlorophytum") || text.includes("dracaena")
      || text.includes("epipremnum") || text.includes("fittonia") || text.includes("nephrolepis")
      || text.includes("schefflera") || text.includes("spathiphyllum") || text.includes("sansevieria")
      || text.includes("sanseveria") || text.includes("zamioculcas")
    ) {
      return "Other non-flowering plant";
    }
    if (
      text.includes("curio") || text.includes("crassula") || text.includes("echeveria")
      || text.includes("succulent") || text.includes("cactus")
    ) {
      return "CITES ge. non-flowering p";
    }
    if (
      text.includes("echeveria") || text.includes("fuchsia") || text.includes("gerbera")
      || text.includes("guzmania") || text.includes("kalanchoe") || text.includes("phalaenopsis")
      || text.includes("anthurium") || text.includes("celosia") || text.includes("cymbidium")
      || text.includes("cyclamen") || text.includes("begonia") || text.includes("helianthus")
      || text.includes("helleanthus") || text.includes("hydrangea") || text.includes("mandevilla")
      || text.includes("lithodora") || text.includes("platycodon") || text.includes("chrysanthem")
      || text.includes("dianthus")
    ) {
      return "Flowering plants(no cactu";
    }
    return knownGroup || "Other non-flowering plant";
  }

  if (code.startsWith("060240") || code.startsWith("60240")) {
    return "refined roses";
  }
  if (code.startsWith("06029091") || code.startsWith("6029091")) {
    return "Flowering plants(no cactu";
  }
  if (code.startsWith("060290500") || code.startsWith("60290500") || code.startsWith("6029050")) {
    return "Perennials";
  }
  if (code.startsWith("060319700") || code.startsWith("60319700") || code === "6031970") {
    return "Others";
  }
  if (
    knownGroup
    && (
      knownGroup === "refined roses"
      || knownGroup === "Perennials"
      || knownGroup === "Others"
      || knownGroup === "Flowering plants(no cactu"
    )
  ) {
    return knownGroup;
  }
  if (
    code.startsWith("060290990") || code.startsWith("60290990")
    || code.startsWith("060290991") || code.startsWith("60290991")
  ) {
    const ambiguousGroup = mapUkdocsCsiAmbiguousPlantGroup(text);
    if (ambiguousGroup) {
      return ambiguousGroup;
    }
    if (knownGroup === "CITES ge. non-flowering p" || knownGroup === "Other non-flowering plant") {
      return knownGroup;
    }
  }

  if (preferPlants) {
    if (text.includes("bonsai") || text.includes("sageretia") || text.includes("aloe") || text.includes("rhipsalis")) {
      return "CITES ge. non-flowering p";
    }
    if (text.includes("salvia") || text.includes("lavandula") || text.includes("helleborus") || text.includes("campanula")) {
      return "Perennials";
    }
    if (text.includes("hibiscus")) {
      return "Flowering plants(no cactu";
    }
    if (text.includes("cupressus")) {
      return "Perennials";
    }
    if (text.includes("ficus")) {
      return "Others";
    }
    if (text.includes("rosa") || (text.includes("rose") && !text.includes("hibiscus"))) {
      return "refined roses";
    }
    if (
      text.includes("dypsis") || text.includes("maranta") || text.includes("calathea")
      || text.includes("chlorophytum") || text.includes("dracaena")
      || text.includes("epipremnum") || text.includes("fittonia") || text.includes("nephrolepis")
      || text.includes("schefflera") || text.includes("spathiphyllum") || text.includes("sansevieria")
      || text.includes("sanseveria") || text.includes("zamioculcas")
    ) {
      return "Other non-flowering plant";
    }
    if (
      text.includes("curio") || text.includes("crassula") || text.includes("echeveria")
      || text.includes("succulent") || text.includes("cactus")
    ) {
      return "CITES ge. non-flowering p";
    }
    if (
      text.includes("echeveria") || text.includes("fuchsia") || text.includes("gerbera")
      || text.includes("guzmania") || text.includes("kalanchoe") || text.includes("phalaenopsis")
      || text.includes("anthurium") || text.includes("celosia") || text.includes("cymbidium")
      || text.includes("cyclamen") || text.includes("begonia") || text.includes("helianthus")
      || text.includes("helleanthus") || text.includes("hydrangea") || text.includes("mandevilla")
      || text.includes("lithodora") || text.includes("platycodon")
    ) {
      return "Flowering plants(no cactu";
    }
    if (text.includes("chrysanthem") || text.includes("dianthus")) {
      return "Flowering plants(no cactu";
    }
  }

  if (text.includes("chrysanthem") || code.startsWith("060314") || code.startsWith("603140")) {
    return "Flowers chrysanthemums";
  }
  if (text.includes("dianthus") || text.includes("carnation") || code.startsWith("060312") || code.startsWith("603120")) {
    return "Flowers carnation";
  }
  if (text.includes("rosa") || text.includes("rose") || code.startsWith("060311") || code.startsWith("603110")) {
    return "Flowers roses";
  }
  if (text.includes("lil") || text.includes("lilium") || code.startsWith("060315") || code.startsWith("603150")) {
    return "Flowers lilies";
  }
  if (text.includes("orchid") || text.includes("dendrob") || code.startsWith("060313") || code.startsWith("603130")) {
    return "Flowers orchids";
  }
  if (text.includes("green") || code.startsWith("0604209") || code.startsWith("604209")) {
    return "Flowers green";
  }
  if (
    text.includes("gypsoph") || text.includes("solidago") || text.includes("other fresh")
    || code.startsWith("0603197") || code.startsWith("603197") || code.startsWith("0603199") || code.startsWith("603199")
  ) {
    return preferPlants ? "Others" : "Flowers (other fresh)";
  }
  return String(description || "").trim() || "Unknown product";
}

function getUkdocsCsiComparisonGroup(row) {
  const comparisonGroup = String(row?.comparison_group || "").trim();
  if (comparisonGroup) {
    return comparisonGroup;
  }
  return String(row?.mapped_product || "").trim();
}

function isUkdocsCsiAggregateProductName(productName) {
  const normalized = normalizeUkdocsCsiToken(productName);
  return normalized.includes("mixed cut flowers and branches")
    || normalized.includes("cutflowers and branches")
    || normalized === "total"
    || normalized === "mixed";
}

function addUkdocsCsiQuantity(map, key, quantity) {
  if (!key) {
    return;
  }
  const numeric = Number(quantity);
  if (!Number.isFinite(numeric)) {
    return;
  }
  map.set(key, (map.get(key) || 0) + numeric);
}

function getUkdocsCsiCurrenciesFromText(text) {
  const source = String(text || "");
  const detected = new Set();
  const codeMatches = source.toUpperCase().match(/\b(?:GBP|EUR|USD)\b/g) || [];
  for (const match of codeMatches) {
    detected.add(match);
  }
  if (source.includes("€")) {
    detected.add("EUR");
  }
  if (source.includes("£")) {
    detected.add("GBP");
  }
  if (source.includes("$")) {
    detected.add("USD");
  }
  return Array.from(detected);
}

function csiStatusRank(status) {
  const normalized = String(status || "").trim().toLowerCase();
  if (normalized === "fail") {
    return 3;
  }
  if (normalized === "warn") {
    return 2;
  }
  if (normalized === "pass") {
    return 1;
  }
  return 0;
}

function mergeUkdocsCsiStatus(...statuses) {
  let best = "";
  let rank = 0;
  for (const status of statuses) {
    const nextRank = csiStatusRank(status);
    if (nextRank > rank) {
      rank = nextRank;
      best = String(status || "").trim().toLowerCase();
    }
  }
  return best || "pass";
}

function uniqueUkdocsCsiStrings(values) {
  return Array.from(new Set((Array.isArray(values) ? values : []).map((value) => String(value || "").trim()).filter(Boolean)));
}

function isUkdocsNoPdNeeded(collection) {
  const pdCodeCompact = String(collection?.pd_code || "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "");
  const pdTypeCompact = String(collection?.pd_type || "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "");
  return pdCodeCompact.includes("nopdneeded")
    || pdCodeCompact === "nopd"
    || pdTypeCompact.includes("nophytoneeded")
    || pdTypeCompact.includes("nopdneeded");
}

function finalizeUkdocsCsiOverallStatus(checks, products) {
  let overall = "pass";
  for (const item of [...(Array.isArray(checks) ? checks : []), ...(Array.isArray(products) ? products : [])]) {
    overall = mergeUkdocsCsiStatus(overall, item?.status || "");
  }
  return overall || "warn";
}

function parseUkdocsCsiReportQuantity(value) {
  const text = String(value || "").trim();
  if (!text) {
    return null;
  }
  const numeric = Number(text);
  return Number.isFinite(numeric) ? numeric : null;
}

function buildUkdocsCsiDomainProducts(baseProducts, sourceRows, domain) {
  const scopedSources = domain === "plants"
    ? {
      ipaffs: new Set(["ipaffs_plants"]),
      tempPhyto: new Set(["temp_phyto_plants", "visual_temp_phyto_plants"]),
    }
    : {
      ipaffs: new Set(["ipaffs"]),
      tempPhyto: new Set(["temp_phyto", "visual_temp_phyto"]),
    };

  return (Array.isArray(baseProducts) ? baseProducts : [])
    .filter((item) => item?.product_domain === domain)
    .map((item) => {
      const productName = String(item?.product || "").trim();
      const invoiceQty = parseUkdocsCsiReportQuantity(item?.invoice_quantity);
      const exportQty = parseUkdocsCsiReportQuantity(item?.export_quantity);
      const ipaffsQty = (Array.isArray(sourceRows) ? sourceRows : [])
        .filter((row) => getUkdocsCsiComparisonGroup(row) === productName && scopedSources.ipaffs.has(String(row?.source || "").trim()))
        .reduce((sum, row) => sum + (Number.isFinite(Number(row?.quantity)) ? Number(row.quantity) : 0), 0);
      const tempPhytoPerDocumentMap = new Map();
      for (const row of Array.isArray(sourceRows) ? sourceRows : []) {
        if (getUkdocsCsiComparisonGroup(row) !== productName) {
          continue;
        }
        if (!scopedSources.tempPhyto.has(String(row?.source || "").trim())) {
          continue;
        }
        const label = String(row?.document_label || "").trim();
        if (!label) {
          continue;
        }
        const quantity = Number(row?.quantity);
        if (!Number.isFinite(quantity)) {
          continue;
        }
        tempPhytoPerDocumentMap.set(label, (tempPhytoPerDocumentMap.get(label) || 0) + quantity);
      }
      const tempPhytoQuantities = Array.from(tempPhytoPerDocumentMap.entries()).map(([documentLabel, quantity]) => ({
        document_label: documentLabel,
        quantity: String(quantity),
      }));
      const tempPhytoQty = tempPhytoQuantities.reduce((sum, entry) => sum + (Number(entry.quantity) || 0), 0);
      const expectedQty = exportQty ?? invoiceQty;
      const messages = [];
      let status = "pass";

      if (invoiceQty !== null || exportQty !== null) {
        if (invoiceQty === null || exportQty === null) {
          status = mergeUkdocsCsiStatus(status, "warn");
          messages.push("Missing in invoice or export file.");
        } else if (invoiceQty !== exportQty) {
          status = mergeUkdocsCsiStatus(status, "warn");
          messages.push(`Invoice/export differ by ${Math.abs(invoiceQty - exportQty)}.`);
        } else {
          messages.push("Invoice/export match.");
        }
      }

      if (ipaffsQty > 0) {
        if (expectedQty === null) {
          status = mergeUkdocsCsiStatus(status, "warn");
          messages.push("IPAFFS has quantity but invoice/export is missing.");
        } else if (ipaffsQty > expectedQty) {
          status = mergeUkdocsCsiStatus(status, "warn");
          messages.push(`IPAFFS quantity ${ipaffsQty} is higher than invoice/export ${expectedQty}.`);
        } else if (ipaffsQty === expectedQty) {
          messages.push("IPAFFS matches invoice/export.");
        } else {
          messages.push(`IPAFFS subset ${ipaffsQty} is within invoice/export ${expectedQty}.`);
        }
      }

      if (tempPhytoQty > 0) {
        if (expectedQty === null) {
          status = mergeUkdocsCsiStatus(status, "warn");
          messages.push(`Temp phyto quantity ${tempPhytoQty} has no matching invoice/export line.`);
        } else if (tempPhytoQty > expectedQty) {
          status = mergeUkdocsCsiStatus(status, "warn");
          messages.push(`Temp phyto quantity ${tempPhytoQty} is higher than invoice/export ${expectedQty}.`);
        } else if (tempPhytoQty === expectedQty) {
          messages.push(`Temp phyto quantity ${tempPhytoQty} matches invoice/export.`);
        } else {
          messages.push(`Temp phyto subset ${tempPhytoQty} is within invoice/export ${expectedQty}.`);
        }
      }

      if (ipaffsQty > 0 && tempPhytoQty > 0) {
        if (ipaffsQty !== tempPhytoQty) {
          status = mergeUkdocsCsiStatus(status, "warn");
          if (ipaffsQty < tempPhytoQty) {
            messages.push(`IPAFFS quantity ${ipaffsQty} is lower than temp phyto ${tempPhytoQty}.`);
          } else {
            messages.push(`IPAFFS quantity ${ipaffsQty} is higher than temp phyto ${tempPhytoQty}.`);
          }
        } else {
          messages.push(`IPAFFS and temp phyto match at ${ipaffsQty}.`);
        }
      }

      return {
        ...item,
        ipaffs_quantity: ipaffsQty > 0 ? String(ipaffsQty) : "",
        temp_phyto_quantity: tempPhytoQty > 0 ? String(tempPhytoQty) : "",
        temp_phyto_quantities: tempPhytoQuantities,
        status,
        message: messages.join(" "),
      };
    });
}

function getUkdocsCsiTempPhytoParsedSourceDocuments(collection) {
  const tempPhytoFiles = (collection?.documents?.temp_phyto_files || []).map((document) => ({
    kind: "temp_phyto",
    prefer_plants: inferUkdocsCsiPlantsPreference(document, false),
    document,
  }));
  const plantTempPhytoDocument = collection?.documents?.temp_phyto_plants_xml_file?.storage_name
    ? collection.documents.temp_phyto_plants_xml_file
    : collection?.documents?.temp_phyto_plants_file?.storage_name
      ? collection.documents.temp_phyto_plants_file
      : null;
  const plantTempPhytoKind = collection?.documents?.temp_phyto_plants_xml_file?.storage_name
    ? "temp_phyto_plants_xml_file"
    : "temp_phyto_plants_file";
  const plantTempPhyto = plantTempPhytoDocument?.storage_name
    ? [{
      kind: plantTempPhytoKind,
      prefer_plants: inferUkdocsCsiPlantsPreference(plantTempPhytoDocument, true),
      document: plantTempPhytoDocument,
    }]
    : [];
  return [...tempPhytoFiles, ...plantTempPhyto];
}

function getUkdocsCsiTempPhytoVisionSourceDocuments(collection) {
  const tempPhytoFiles = (collection?.documents?.temp_phyto_files || []).map((document) => ({
    kind: "temp_phyto",
    prefer_plants: inferUkdocsCsiPlantsPreference(document, false),
    document,
  }));
  const plantTempPhyto = (!collection?.documents?.temp_phyto_plants_xml_file?.storage_name && collection?.documents?.temp_phyto_plants_file?.storage_name)
    ? [{
      kind: "temp_phyto_plants_file",
      prefer_plants: inferUkdocsCsiPlantsPreference(collection.documents.temp_phyto_plants_file, true),
      document: collection.documents.temp_phyto_plants_file,
    }]
    : [];
  return [...tempPhytoFiles, ...plantTempPhyto];
}

function isUkdocsCsiTempPhytoDeterministicReady(parsedData) {
  const parsed = parsedData && typeof parsedData === "object" ? parsedData : {};
  const sourceFormat = String(parsed?.source_format || "").trim().toLowerCase();
  const productLines = Array.isArray(parsed?.product_lines) ? parsed.product_lines : [];
  const hasProductLines = productLines.some((line) => String(line?.product || "").trim() && Number.isFinite(Number(line?.quantity)));
  const hasPcnu = Boolean(String(parsed?.pcnu_number || "").trim());
  const state = String(parsed?.document_state || "").trim().toLowerCase();
  const problems = Array.isArray(parsed?.problems) ? parsed.problems.filter(Boolean) : [];
  const validationReport = parsed?.validation_report && typeof parsed.validation_report === "object"
    ? parsed.validation_report
    : null;
  const validationLooksGood = !validationReport
    || (
      Number(validationReport?.duplicate_count || 0) === 0
      && Number(validationReport?.missing_quantity_count || 0) === 0
      && Number(validationReport?.extracted_row_count || 0) > 0
    );

  if (!hasProductLines || !hasPcnu || state === "not_activated") {
    return false;
  }
  if (sourceFormat === "xml") {
    return validationLooksGood;
  }
  return problems.length === 0;
}

function normalizeUkdocsCsiDocumentName(value) {
  return String(value || "").trim().toLowerCase();
}

function normalizeUkdocsCsiDocumentLabel(value) {
  return String(value || "").trim().toLowerCase();
}

function buildUkdocsCsiDeterministicReport(collection, extractedDocuments) {
  const documents = Array.isArray(extractedDocuments) ? extractedDocuments : [];
  const noPdNeeded = isUkdocsNoPdNeeded(collection);
  const invoiceDocs = documents.filter((document) => document?.kind === "generated_invoice");
  const exportDoc = documents.find((document) => document?.kind === "generated_export") || null;
  const storedIpaffsDocs = ["ipaffs_file", "ipaffs_plants_file"]
    .map((kind) => {
      const source = collection?.documents?.[kind];
      return source?.parsed_data
        ? {
          kind,
          name: String(source.original_name || source.storage_name || "").trim(),
          content_type: String(source.content_type || source.mime_type || "").trim(),
          line_count: Number(source.line_count || 0),
          delimiter: String(source.delimiter || source?.parsed_data?.delimiter || "").trim(),
          parsed_data: source.parsed_data,
          error: String(source.parse_error || "").trim(),
        }
        : null;
    })
    .filter(Boolean);
  const rawIpaffsDocs = (() => {
    const extracted = documents.filter((document) => document?.kind === "ipaffs_file" || document?.kind === "ipaffs_plants_file");
    const merged = [...extracted];
    for (const storedDocument of storedIpaffsDocs) {
      if (!merged.some((item) => String(item?.kind || "").trim() === String(storedDocument.kind || "").trim())) {
        merged.push(storedDocument);
      }
    }
    return merged;
  })();
  const rawTempPhytoDocs = (() => {
    const sourceDocuments = getUkdocsCsiTempPhytoParsedSourceDocuments(collection);
    const extractedTempDocs = documents.filter((document) => document?.kind === "temp_phyto" || document?.kind === "temp_phyto_plants_file" || document?.kind === "temp_phyto_plants_xml_file");
    const sourceDocumentByKey = new Map(
      sourceDocuments.map((item, index) => {
        const key = `${String(item?.kind || "").trim()}::${normalizeUkdocsCsiDocumentName(item?.document?.original_name || item?.document?.storage_name)}`;
        return [key || `source-index-${index}`, item];
      }),
    );
    if (extractedTempDocs.length) {
      const extractedDocByKey = new Map(
        extractedTempDocs.map((document, index) => {
          const documentKey = `${String(document?.kind || "").trim()}::${normalizeUkdocsCsiDocumentName(document?.name)}`;
          return [documentKey || `extracted-index-${index}`, document];
        }),
      );
      return sourceDocuments.map((sourceItem, index) => {
        const fallback = sourceItem?.document || null;
        const sourceKey = `${String(sourceItem?.kind || "").trim()}::${normalizeUkdocsCsiDocumentName(fallback?.original_name || fallback?.storage_name)}`;
        const document = extractedDocByKey.get(sourceKey) || extractedTempDocs[index] || null;
        if (!document) {
          return {
            kind: sourceItem?.kind || "",
            prefer_plants: sourceItem?.prefer_plants === true,
            name: String(fallback?.original_name || fallback?.storage_name || "").trim(),
            content_type: String(fallback?.content_type || fallback?.mime_type || "").trim(),
            line_count: Number(fallback?.line_count || 0),
            parsed_data: fallback?.parsed_data || null,
            error: String(fallback?.parse_error || "").trim(),
          };
        }
        if (!fallback?.parsed_data) {
          return {
            ...document,
            prefer_plants: sourceItem.prefer_plants === true,
          };
        }
        const extractedHasLines = Array.isArray(document?.parsed_data?.product_lines) && document.parsed_data.product_lines.length > 0;
        const extractedHasPcnu = String(document?.parsed_data?.pcnu_number || "").trim();
        if (extractedHasLines && extractedHasPcnu) {
          return {
            ...document,
            prefer_plants: sourceItem.prefer_plants === true,
          };
        }
        return {
          ...document,
          prefer_plants: sourceItem.prefer_plants === true,
          line_count: Number(document?.line_count || fallback.line_count || 0),
          parsed_data: fallback.parsed_data,
          error: String(document?.error || fallback.parse_error || "").trim(),
        };
      });
    }
    return sourceDocuments
      .filter((item) => item?.document?.parsed_data)
      .map((item) => ({
        kind: item.kind,
        prefer_plants: item.prefer_plants === true,
        name: String(item.document.original_name || item.document.storage_name || "").trim(),
        content_type: String(item.document.content_type || item.document.mime_type || "").trim(),
        line_count: Number(item.document.line_count || 0),
        parsed_data: item.document.parsed_data,
        error: String(item.document.parse_error || "").trim(),
      }));
  })();
  const ipaffsDocs = noPdNeeded ? [] : rawIpaffsDocs;
  const tempPhytoDocs = noPdNeeded ? [] : rawTempPhytoDocs;
  const hasIpaffsAttached = noPdNeeded
    ? false
    : Boolean(collection?.documents?.ipaffs_file?.storage_name || collection?.documents?.ipaffs_plants_file?.storage_name);

  const invoiceTotals = new Map();
  const exportTotals = new Map();
  const ipaffsTotals = new Map();
  const tempPhytoTotals = new Map();
  const products = [];
  const sourceRows = [];
  const checks = [];
  const notes = [];
  const manualChecks = [];
  const extractedKinds = documents.map((document) => String(document?.kind || "").trim()).filter(Boolean);
  const extractedIpaffsDebug = ipaffsDocs.map((ipaffsDoc) => ({
    kind: String(ipaffsDoc?.kind || "").trim(),
    name: String(ipaffsDoc?.name || "").trim(),
    content_type: String(ipaffsDoc?.content_type || "").trim(),
    line_count: Number.isFinite(Number(ipaffsDoc?.line_count)) ? Number(ipaffsDoc.line_count) : 0,
    has_parsed_data: Boolean(ipaffsDoc?.parsed_data && typeof ipaffsDoc.parsed_data === "object"),
    row_count: Array.isArray(ipaffsDoc?.parsed_data?.rows) ? ipaffsDoc.parsed_data.rows.length : 0,
    delimiter: String(ipaffsDoc?.parsed_data?.delimiter || ipaffsDoc?.delimiter || "").trim(),
    error: String(ipaffsDoc?.error || "").trim(),
  }));

  for (const document of invoiceDocs) {
    const rows = Array.isArray(document?.parsed_data?.rows) ? document.parsed_data.rows : [];
    const strictDomain = inferUkdocsCsiDocumentDomain(document?.name || "", document?.kind || "");
    for (const row of rows) {
      const rawProduct = String(row?.product || "").trim();
      const comparisonGroup = mapUkdocsCsiProductName(rawProduct, row?.commodity_code, {
        ...(strictDomain ? { strict_domain: strictDomain } : {}),
        document_name: document?.name || "",
      });
      addUkdocsCsiQuantity(
        invoiceTotals,
        comparisonGroup,
        row?.quantity,
      );
      sourceRows.push({
        source: "invoice",
        document_name: String(document?.name || "").trim(),
        raw_product: rawProduct,
        commodity_code: String(row?.commodity_code || "").trim(),
        mapped_product: rawProduct || comparisonGroup,
        comparison_group: comparisonGroup,
        product_domain: getUkdocsCsiProductDomain(comparisonGroup),
        quantity: Number.isFinite(Number(row?.quantity)) ? Number(row.quantity) : null,
      });
    }
  }

  const exportRows = Array.isArray(exportDoc?.parsed_data?.rows) ? exportDoc.parsed_data.rows : [];
  const exportStrictDomain = inferUkdocsCsiDocumentDomain(exportDoc?.name || "", exportDoc?.kind || "");
  for (const row of exportRows) {
    const rawProduct = String(row?.product || "").trim();
    const comparisonGroup = mapUkdocsCsiProductName(rawProduct, row?.commodity_code, {
      ...(exportStrictDomain ? { strict_domain: exportStrictDomain } : {}),
      document_name: exportDoc?.name || "",
    });
    addUkdocsCsiQuantity(
      exportTotals,
      comparisonGroup,
      row?.quantity,
    );
    sourceRows.push({
      source: "export",
      document_name: String(exportDoc?.name || "").trim(),
      raw_product: rawProduct,
      commodity_code: String(row?.commodity_code || "").trim(),
      mapped_product: rawProduct || comparisonGroup,
      comparison_group: comparisonGroup,
      product_domain: getUkdocsCsiProductDomain(comparisonGroup),
      quantity: Number.isFinite(Number(row?.quantity)) ? Number(row.quantity) : null,
    });
  }

  const ipaffsRows = ipaffsDocs.flatMap((ipaffsDoc) => (Array.isArray(ipaffsDoc?.parsed_data?.rows) ? ipaffsDoc.parsed_data.rows : []));
  for (const [ipaffsIndex, row] of ipaffsRows.entries()) {
    const sourceDoc = ipaffsDocs.find((item) => {
      const rows = Array.isArray(item?.parsed_data?.rows) ? item.parsed_data.rows : [];
      return rows.includes(row);
    }) || null;
    const mappedProduct = String(row?.mapped_group || "").trim() || mapUkdocsCsiProductName(row?.product || row?.genus, row?.commodity_code, {
      prefer_plants: sourceDoc?.kind === "ipaffs_plants_file",
      document_name: sourceDoc?.name || "",
      strict_domain: sourceDoc?.kind === "ipaffs_plants_file" ? "plants" : "flowers",
    });
    addUkdocsCsiQuantity(
      ipaffsTotals,
      mappedProduct,
      row?.quantity,
    );
    sourceRows.push({
      source: sourceDoc?.kind === "ipaffs_plants_file" ? "ipaffs_plants" : "ipaffs",
      document_name: String(sourceDoc?.name || `IPAFFS row ${ipaffsIndex + 1}`).trim(),
      raw_product: String(row?.genus || row?.product || "").trim(),
      commodity_code: String(row?.commodity_code || "").trim(),
      mapped_product: mappedProduct,
      comparison_group: mappedProduct,
      product_domain: getUkdocsCsiProductDomain(mappedProduct),
      quantity: Number.isFinite(Number(row?.quantity)) ? Number(row.quantity) : null,
    });
  }

  const tempPhytoContexts = tempPhytoDocs.map((document, index) => {
    const parsed = document?.parsed_data && typeof document.parsed_data === "object" ? document.parsed_data : {};
    const lineProducts = [];
    const documentLabel = `Temp phyto ${String.fromCharCode(65 + index)}`;
    for (const line of Array.isArray(parsed?.product_lines) ? parsed.product_lines : []) {
      const mappedProduct = mapUkdocsCsiProductName(line?.product || "", "", {
        document_name: document?.name || "",
        prefer_plants: document?.prefer_plants === true,
        strict_domain: document?.prefer_plants === true ? "plants" : "flowers",
      });
      addUkdocsCsiQuantity(tempPhytoTotals, mappedProduct, line?.quantity);
      lineProducts.push({
        product: mappedProduct,
        raw_product: String(line?.product || "").trim(),
        quantity: Number.isFinite(Number(line?.quantity)) ? Number(line.quantity) : null,
      });
      sourceRows.push({
        source: document?.prefer_plants === true ? "temp_phyto_plants" : "temp_phyto",
        document_name: String(document?.name || documentLabel).trim(),
        document_label: documentLabel,
        raw_product: String(line?.product || "").trim(),
        commodity_code: "",
        mapped_product: mappedProduct,
        comparison_group: mappedProduct,
        product_domain: getUkdocsCsiProductDomain(mappedProduct),
        quantity: Number.isFinite(Number(line?.quantity)) ? Number(line.quantity) : null,
      });
    }
    const problems = Array.isArray(parsed?.problems) ? parsed.problems : [];
    if (parsed?.document_state === "not_activated") {
      checks.push({
        code: "TEMP_PHYTO_STATE",
        status: "fail",
        message: `${document.name || "Temporary phyto PDF"} appears not activated.`,
      });
    }
    if (!parsed?.pcnu_number) {
      manualChecks.push(`Check PCNU manually in ${document.name || "temporary phyto PDF"} because no PCNU number was parsed.`);
    }
    if (problems.length) {
      manualChecks.push(`Review ${document.name || "temporary phyto PDF"}: ${problems.join(", ")}.`);
    }
    return {
      document_label: documentLabel,
      name: String(document?.name || "").trim(),
      source_format: String(parsed?.source_format || "").trim().toLowerCase(),
      parsed_pcnu_number: String(parsed?.pcnu_number || "").trim(),
      parsed_document_state: String(parsed?.document_state || "").trim() || "unknown",
      prefer_plants: document?.prefer_plants === true,
      parsed_total_quantity: Number.isFinite(Number(parsed?.total_quantity)) ? Number(parsed.total_quantity) : null,
      deterministic_ready: isUkdocsCsiTempPhytoDeterministicReady(parsed),
      expected_products: lineProducts.map((item) => ({
        product: item.product,
        raw_product: item.raw_product,
        expected_quantity: exportTotals.get(item.product) ?? invoiceTotals.get(item.product) ?? null,
        parsed_quantity: item.quantity,
      })),
      parsed_problems: problems,
    };
  });

  const allProducts = new Set([
    ...invoiceTotals.keys(),
    ...exportTotals.keys(),
    ...ipaffsTotals.keys(),
    ...tempPhytoTotals.keys(),
  ]);

  let invoiceExportMismatchCount = 0;
  let ipaffsMismatchCount = 0;
  let tempPhytoMismatchCount = 0;

  for (const product of Array.from(allProducts).sort()) {
    const invoiceQty = invoiceTotals.has(product) ? invoiceTotals.get(product) : null;
    const exportQty = exportTotals.has(product) ? exportTotals.get(product) : null;
    const ipaffsQty = ipaffsTotals.has(product) ? ipaffsTotals.get(product) : null;
    const phytoQty = tempPhytoTotals.has(product) ? tempPhytoTotals.get(product) : null;
    const phytoPerDocument = tempPhytoContexts
      .map((context) => {
        const match = (context.expected_products || []).find((item) => item.product === product);
        return {
          document_label: context.document_label || context.name || "Temp phyto",
          quantity: match?.parsed_quantity ?? null,
        };
      })
      .filter((item) => item.quantity !== null);
    const expectedQty = exportQty ?? invoiceQty;
    const messages = [];
    let status = "pass";

    if (invoiceQty !== null || exportQty !== null) {
      if (invoiceQty === null || exportQty === null) {
        status = mergeUkdocsCsiStatus(status, "warn");
        invoiceExportMismatchCount += 1;
        messages.push("Missing in invoice or export file.");
      } else if (invoiceQty !== exportQty) {
        status = mergeUkdocsCsiStatus(status, "warn");
        invoiceExportMismatchCount += 1;
        messages.push(`Invoice/export differ by ${Math.abs(invoiceQty - exportQty)}.`);
      } else {
        messages.push("Invoice/export match.");
      }
    }

    if (ipaffsDocs.length) {
      if (ipaffsQty !== null) {
        if (expectedQty === null) {
          status = mergeUkdocsCsiStatus(status, "warn");
          ipaffsMismatchCount += 1;
          messages.push("IPAFFS has quantity but invoice/export is missing.");
        } else if (ipaffsQty > expectedQty) {
          status = mergeUkdocsCsiStatus(status, "warn");
          ipaffsMismatchCount += 1;
          messages.push(`IPAFFS quantity ${ipaffsQty} is higher than invoice/export ${expectedQty}.`);
        } else if (ipaffsQty === expectedQty) {
          messages.push("IPAFFS matches invoice/export.");
        } else {
          messages.push(`IPAFFS subset ${ipaffsQty} is within invoice/export ${expectedQty}.`);
        }
      }
    }

    if (phytoQty !== null) {
      if (expectedQty === null) {
        status = mergeUkdocsCsiStatus(status, "warn");
        tempPhytoMismatchCount += 1;
        messages.push(`Temp phyto quantity ${phytoQty} has no matching invoice/export line.`);
      } else if (phytoQty > expectedQty) {
        status = mergeUkdocsCsiStatus(status, "warn");
        tempPhytoMismatchCount += 1;
        messages.push(`Temp phyto quantity ${phytoQty} is higher than invoice/export ${expectedQty}.`);
      } else if (phytoQty === expectedQty) {
        messages.push(`Temp phyto quantity ${phytoQty} matches invoice/export.`);
      } else {
        messages.push(`Temp phyto subset ${phytoQty} is within invoice/export ${expectedQty}.`);
      }
    }

    if (ipaffsQty !== null && phytoQty !== null) {
      if (ipaffsQty !== phytoQty) {
        status = mergeUkdocsCsiStatus(status, "warn");
        ipaffsMismatchCount += 1;
        tempPhytoMismatchCount += 1;
        if (ipaffsQty < phytoQty) {
          messages.push(`IPAFFS quantity ${ipaffsQty} is lower than temp phyto ${phytoQty}.`);
        } else {
          messages.push(`IPAFFS quantity ${ipaffsQty} is higher than temp phyto ${phytoQty}.`);
        }
      } else {
        messages.push(`IPAFFS and temp phyto match at ${ipaffsQty}.`);
      }
    }

    products.push({
      product,
      product_domain: getUkdocsCsiProductDomain(product),
      invoice_quantity: invoiceQty === null ? "" : String(invoiceQty),
      export_quantity: exportQty === null ? "" : String(exportQty),
      ipaffs_quantity: ipaffsQty === null ? "" : String(ipaffsQty),
      temp_phyto_quantity: phytoQty === null ? "" : String(phytoQty),
      temp_phyto_quantities: phytoPerDocument.map((item) => ({
        document_label: item.document_label,
        quantity: String(item.quantity),
      })),
      status,
      message: messages.join(" "),
    });
  }

  checks.push({
    code: "INV_EXP_RECONCILIATION",
    status: invoiceExportMismatchCount ? "warn" : "pass",
    message: invoiceExportMismatchCount
      ? `${invoiceExportMismatchCount} invoice/export product totals need review.`
      : "Combined invoice quantities match generated export quantities.",
  });

  checks.push({
    code: "IPAFFS_VERIFICATION",
    status: noPdNeeded
      ? "pass"
      : !hasIpaffsAttached
      ? "warn"
      : !ipaffsDocs.length
        ? "warn"
        : extractedIpaffsDebug.some((item) => item.error)
          ? "warn"
        : !ipaffsRows.length
          ? "warn"
          : ipaffsMismatchCount
            ? "warn"
            : "pass",
    message: noPdNeeded
      ? "IPAFFS is not required because PD code indicates no PD needed."
      : !hasIpaffsAttached
      ? "IPAFFS file is missing."
      : !ipaffsDocs.length
        ? `IPAFFS file is attached on the zending, but CSI extraction returned these document kinds only: ${extractedKinds.join(", ") || "none"}.`
        : extractedIpaffsDebug.some((item) => item.error)
          ? `One or more IPAFFS files could not be extracted: ${extractedIpaffsDebug.filter((item) => item.error).map((item) => `${item.name || item.kind}: ${item.error}`).join("; ")}`
        : !ipaffsRows.length
          ? `IPAFFS file(s) were loaded (${ipaffsDocs.map((item) => item.name || item.kind || "unknown file").join(", ")}), but no product rows were parsed.`
          : ipaffsMismatchCount
            ? `${ipaffsMismatchCount} IPAFFS or IPAFFS/temp phyto product totals need review.`
            : "IPAFFS product totals stay within invoice/export and match temp phyto where both exist.",
  });

  checks.push({
    code: "TEMP_PHYTO_FILES",
    status: noPdNeeded ? "pass" : (tempPhytoDocs.length ? "pass" : "warn"),
    message: noPdNeeded
      ? "Temporary phyto PDF files are not required because PD code indicates no PD needed."
      : tempPhytoDocs.length
      ? `${tempPhytoDocs.length} temporary phyto PDF file(s) attached.`
      : "No temporary phyto PDF files attached.",
  });

  if (tempPhytoDocs.length) {
    const hasActiveProblems = tempPhytoContexts.some((item) => item.parsed_document_state === "not_activated");
    const missingPcnuCount = tempPhytoContexts.filter((item) => !item.parsed_pcnu_number).length;
    checks.push({
      code: "TEMP_PHYTO_PARSE",
      status: hasActiveProblems
        ? "fail"
        : missingPcnuCount
          ? "warn"
          : tempPhytoMismatchCount
            ? "warn"
            : "pass",
      message: hasActiveProblems
        ? "At least one temporary phyto PDF is not activated."
        : missingPcnuCount
          ? `${missingPcnuCount} temporary phyto PDF(s) are missing a parsed PCNU number from PDF text extraction.`
          : tempPhytoMismatchCount
            ? `${tempPhytoMismatchCount} temp phyto or temp phyto/IPAFFS quantity comparison(s) need review.`
            : "Temporary phyto quantities stay within invoice/export and match IPAFFS where both exist.",
    });
  }

  const exportCurrencies = getUkdocsCsiCurrenciesFromText(exportDoc?.text || "");
  const invoiceCurrencies = Array.from(new Set(invoiceDocs.flatMap((document) => getUkdocsCsiCurrenciesFromText(document?.text || ""))));
  const combinedCurrencies = Array.from(new Set([...exportCurrencies, ...invoiceCurrencies]));
  checks.push({
    code: "CURRENCY_CHECK",
    status: combinedCurrencies.length > 1 ? "warn" : "pass",
    message: combinedCurrencies.length > 1
      ? `Multiple currencies found in extracted text: ${combinedCurrencies.join(", ")}.`
      : combinedCurrencies.length === 1
        ? `Currency appears consistent as ${combinedCurrencies[0]}.`
        : "No explicit currency code found in extracted text; keep a quick visual check.",
  });
  if (!combinedCurrencies.length) {
    manualChecks.push("Confirm the invoice/export currency manually because no explicit currency code was extracted.");
  }

  const flowerProducts = buildUkdocsCsiDomainProducts(products, sourceRows, "flowers");
  const plantProducts = buildUkdocsCsiDomainProducts(products, sourceRows, "plants");
  const combinedProducts = [...flowerProducts, ...plantProducts].sort((left, right) => (
    String(left?.product || "").localeCompare(String(right?.product || ""))
  ));
  const overallStatus = finalizeUkdocsCsiOverallStatus(checks, combinedProducts);
  const summaryParts = [
    invoiceExportMismatchCount
      ? `${invoiceExportMismatchCount} invoice/export mismatch(es)`
      : "invoice/export quantities match",
    noPdNeeded
      ? "no PD documents needed"
      : !hasIpaffsAttached
      ? "IPAFFS missing"
      : !ipaffsDocs.length
        ? "IPAFFS extraction missing"
        : !ipaffsRows.length
          ? "IPAFFS rows not parsed"
          : (ipaffsMismatchCount ? `${ipaffsMismatchCount} IPAFFS mismatch(es)` : "IPAFFS matches"),
    noPdNeeded
      ? "temp phyto not needed"
      : tempPhytoDocs.length
      ? (tempPhytoMismatchCount ? `${tempPhytoMismatchCount} temp phyto mismatch(es)` : "temp phyto quantities checked")
      : "no temp phyto PDFs",
  ];

  return {
    report: {
      overall_status: overallStatus,
      summary: summaryParts.join(", ") + ".",
      checks,
      products: combinedProducts,
      flower_products: flowerProducts,
      plant_products: plantProducts,
      source_rows: sourceRows,
      manual_checks: uniqueUkdocsCsiStrings(manualChecks),
      notes: uniqueUkdocsCsiStrings([
        ...notes,
        ...(extractedIpaffsDebug.length
          ? extractedIpaffsDebug.map((item) => `IPAFFS extractor debug: kind=${item.kind || "-"}, name=${item.name || "-"}, content_type=${item.content_type || "-"}, lines=${item.line_count}, rows=${item.row_count}, delimiter=${item.delimiter || "-"}${item.error ? `, error=${item.error}` : ""}.`)
          : (hasIpaffsAttached ? ["IPAFFS extractor debug: attached IPAFFS file was not returned by CSI extraction."] : [])),
      ]),
    },
    visual_context: {
      temp_phyto_documents: tempPhytoContexts,
    },
  };
}

function buildUkdocsCsiAuditPayload(collection, deterministicBundle, requestUser, options = {}) {
  const visionDocuments = Array.isArray(options.vision_documents) ? options.vision_documents.filter(Boolean) : [];
  const visualContext = options.visual_context && typeof options.visual_context === "object"
    ? options.visual_context
    : deterministicBundle?.visual_context;
  const tempPhytoFiles = visionDocuments.length
    ? visionDocuments.map((file) => file.name || "temp-phyto.pdf")
    : (collection?.documents?.temp_phyto_files || []).map((file) => file.original_name || file.storage_name);
  const prompt = {
    task: "UKDocs CSI temporary phyto visual verification",
    instructions: [
      "Return strict JSON only.",
      "Use only the temporary phyto PDF page images for this task.",
      "Check only: visible PCNU number, blocked or not activated state, and clearly visible individual product-line Pieces quantities.",
      "Use expected_temp_phyto_checks for this exact document as your hunt list and match visible lines one by one.",
      "Use raw_product as the visible alias and product as the final CSI group.",
      "If expected_products is empty, you may still return clearly visible individual product lines using the best obvious group for that document domain.",
      "For flower documents, keep flower groups. For plant documents, keep plant groups. Do not switch domains.",
      "Ignore consignee text, invoice totals, export totals, IPAFFS totals, legal text, annex text, and repeated continuation echoes.",
      "If a document has multiple visible product lines, return separate visible_documents rows.",
      "If a value is not clearly visible, do not guess.",
      "If the document looks active and the PCNU number is readable, say so directly.",
      "Never use the Packages or Box count as the product quantity; use only the Pieces quantity.",
      "Never use a page total as a product quantity.",
      "Never combine multiple visible product lines into one quantity.",
      "Never invent a quantity that is not printed on the page.",
      "Never emit a quantity of 1 unless the page clearly shows '1 Pieces' for that same product line.",
      "Never output explanation text before or after the JSON.",
    ],
    output_schema: {
      overall_status: "pass|warn|fail",
      summary: "short human summary",
      checks: [{ code: "string", status: "pass|warn|fail", message: "string" }],
      visible_documents: [{
        document_label: "temp phyto A|temp phyto B",
        product: "expected CSI product/group for the matched visible line",
        quantity: 0,
        pcnu_number: "visible PCNU if readable",
        state: "ok|not_activated|unclear",
        note: "short note",
      }],
      manual_checks: ["string"],
      notes: ["string"],
    },
    zending: {
      shipment_reference: collection?.shipment_reference || "",
      shipment_date: collection?.shipment_date || "",
      customer_name: collection?.customer_name || "",
      pd_type: collection?.pd_type || "",
    },
    temp_phyto_files: tempPhytoFiles,
    expected_temp_phyto_checks: visualContext?.temp_phyto_documents || [],
    return_example: {
      overall_status: "warn",
      summary: "PCNU numbers visible. Individual visible temp phyto product lines were listed where readable.",
      checks: [
        { code: "PHYTO_PCNU_VISIBLE", status: "pass", message: "PCNU 123456789 is visible in temp phyto A." },
        { code: "PHYTO_STATE", status: "pass", message: "No blocked or not activated text is visible." },
      ],
      visible_documents: [
        { document_label: "temp phyto A", product: "Flowers chrysanthemums", quantity: 17160, pcnu_number: "123456789", state: "ok", note: "Matched expected product chrysanthemum to a clearly visible Pieces line." },
        { document_label: "temp phyto A", product: "Flowers (other fresh)", quantity: 25, pcnu_number: "123456789", state: "ok", note: "Matched expected product solidago to a clearly visible Pieces line." },
        { document_label: "temp phyto B", product: "Flowers carnations", quantity: 700, pcnu_number: "987654321", state: "ok", note: "Matched expected product dianthus to a clearly visible Pieces line." },
      ],
      manual_checks: [],
      notes: [
        "Only temporary phyto PDF pages were checked visually.",
      ],
    },
    requested_by: requestUser?.username || "",
  };

  return {
    model: "",
    messages: [
      {
        role: "system",
        content: "You visually verify temporary phytosanitary PDF pages. Return JSON only and never add markdown. Use the provided expected product list as a strict hunt list for matching visible product lines.",
      },
      {
        role: "user",
        content: JSON.stringify(prompt),
      },
    ],
    format: "json",
    think: false,
    options: {
      temperature: 0,
      num_predict: 1800,
    },
  };
}

async function updateUkdocsCsiReport(collectionId, patch) {
  const state = await readUkdocsState();
  const existingCollection = ukdocsPrintCollectionById(state.print_collections, collectionId);
  if (!existingCollection) {
    return null;
  }
  const updatedCollection = normalizeUkdocsPrintCollection({
    ...existingCollection,
    updated_at: new Date().toISOString(),
    csi_report: {
      ...(existingCollection.csi_report || {}),
      ...(patch || {}),
    },
  });
  state.print_collections = upsertUkdocsPrintCollection(state.print_collections, updatedCollection);
  await writeUkdocsState(state);
  return updatedCollection;
}

async function updateUkdocsCsiEmailResult(collectionId, csiEmailPatch) {
  const state = await readUkdocsState();
  const existingCollection = ukdocsPrintCollectionById(state.print_collections, collectionId);
  if (!existingCollection) {
    return null;
  }
  const updatedCollection = normalizeUkdocsPrintCollection({
    ...existingCollection,
    updated_at: new Date().toISOString(),
    csi_email: {
      ...(existingCollection.csi_email || {}),
      ...(csiEmailPatch || {}),
    },
  });
  state.print_collections = upsertUkdocsPrintCollection(state.print_collections, updatedCollection);
  await writeUkdocsState(state);
  return updatedCollection;
}

async function getUkdocsCsiGroupJobs(collectionId, groupId) {
  if (!collectionId || !groupId) {
    return [];
  }
  const result = await dbQuery(
    `
      SELECT *
      FROM llm_jobs
      WHERE collection_id = $1
        AND job_type = 'ukdocs_csi_audit'
        AND payload_json->>'csi_group_id' = $2
      ORDER BY created_at ASC
    `,
    [String(collectionId || "").trim(), String(groupId || "").trim()],
  );
  return Array.isArray(result.rows) ? result.rows.map((row) => ({
    id: String(row?.id || "").trim(),
    job_type: String(row?.job_type || "").trim(),
    status: String(row?.status || "").trim() || "pending",
    created_by: String(row?.created_by || "").trim(),
    shipment_id: String(row?.shipment_id || "").trim(),
    collection_id: String(row?.collection_id || "").trim(),
    document_kind: String(row?.document_kind || "").trim(),
    priority: Number(row?.priority || 0),
    attempt_count: Number(row?.attempt_count || 0),
    max_attempts: Number(row?.max_attempts || 1),
    agent_name: String(row?.agent_name || "").trim(),
    payload_json: typeof row?.payload_json === "object" && row.payload_json ? row.payload_json : {},
    result_json: typeof row?.result_json === "object" && row.result_json ? row.result_json : {},
    error_text: String(row?.error_text || "").trim(),
    created_at: row?.created_at ? new Date(row.created_at).toISOString() : "",
    claimed_at: row?.claimed_at ? new Date(row.claimed_at).toISOString() : "",
    finished_at: row?.finished_at ? new Date(row.finished_at).toISOString() : "",
    updated_at: row?.updated_at ? new Date(row.updated_at).toISOString() : "",
  })) : [];
}

function parseUkdocsCsiAuditJobResult(job) {
  const deterministicSource = normalizeUkdocsCsiParsedResult(job?.payload_json?.deterministic_report) || {
    overall_status: "warn",
    summary: "",
    checks: [],
    products: [],
    flower_products: [],
    plant_products: [],
    source_rows: [],
    manual_checks: [],
    notes: [],
  };
  const contentText = String(job?.result_json?.ollama_response?.message?.content || job?.result_json?.response || "").trim();
  const thinkingText = String(job?.result_json?.ollama_response?.message?.thinking || "").trim();
  const doneReason = String(job?.result_json?.ollama_response?.done_reason || "").trim();
  const llmContent = contentText || thinkingText;
  const parseCandidates = [
    ["result_json.parsed_result", job?.result_json?.parsed_result],
    ["result_json.result", job?.result_json?.result],
    ["result_json.response_json", job?.result_json?.response_json],
    ["ollama_response.message.content", extractJsonObjectFromText(contentText)],
    ["ollama_response.message.thinking", extractJsonObjectFromText(thinkingText)],
    ["result_json.root", job?.result_json],
  ];
  let parseSource = "";
  let parsed = null;
  for (const [sourceName, sourceValue] of parseCandidates) {
    const normalized = normalizeUkdocsCsiParsedResult(sourceValue);
    if (normalized) {
      parseSource = sourceName;
      parsed = normalized;
      break;
    }
  }
  if (!parsed) {
    parsed = {
      overall_status: "warn",
      summary: "",
      checks: [],
      products: [],
      flower_products: [],
      plant_products: [],
      source_rows: [],
      manual_checks: [],
      notes: [],
    };
  }
  const hasStructuredRows = parsed.checks.length
    || parsed.products.length
    || parsed.manual_checks.length
    || (Array.isArray(parsed.visible_documents) && parsed.visible_documents.length);
  const documentLabel = String(job?.payload_json?.csi_document_label || "").trim();
  const labelPrefix = documentLabel ? `${documentLabel}: ` : "";
  const parseError = hasStructuredRows
    ? ""
    : contentText
      ? `${labelPrefix}No structured CSI rows parsed. Source used: ${parseSource || "none"}.`
      : thinkingText
        ? `${labelPrefix}Model returned no final JSON content. It only returned thinking text${doneReason ? ` and stopped with done_reason: ${doneReason}` : ""}. prompt_eval_count=${job?.result_json?.ollama_response?.prompt_eval_count ?? "?"}, eval_count=${job?.result_json?.ollama_response?.eval_count ?? "?"}, thinking_chars=${thinkingText.length}.`
        : `${labelPrefix}No structured CSI rows parsed. Source used: ${parseSource || "none"}.`;
  const llmChecks = hasStructuredRows
    ? (Array.isArray(parsed.checks) ? parsed.checks : [])
    : parseError
      ? [{ code: "LLM_OUTPUT", status: "warn", message: parseError }]
      : [];
  return {
    job,
    deterministicSource,
    contentText,
    thinkingText,
    doneReason,
    llmContent,
    parseSource,
    parsed,
    hasStructuredRows,
    parseError,
    llmChecks,
  };
}

function buildUkdocsCsiReportFromJobResults(jobResults) {
  const results = Array.isArray(jobResults) ? jobResults : [];
  const firstResult = results[0] || null;
  const deterministicSource = firstResult?.deterministicSource || {
    overall_status: "warn",
    summary: "",
    checks: [],
    products: [],
    source_rows: [],
    manual_checks: [],
    notes: [],
  };
  const combinedParsed = {
    summary: uniqueUkdocsCsiStrings(results.map((item) => {
      const summary = String(item?.parsed?.summary || "").trim();
      const label = String(item?.job?.payload_json?.csi_document_label || "").trim();
      return summary ? (label ? `${label}: ${summary}` : summary) : "";
    })).join(" "),
    checks: results.flatMap((item) => item?.llmChecks || []),
    visible_documents: results.flatMap((item) => {
      const canonicalLabel = String(item?.job?.payload_json?.csi_document_label || "").trim();
      return Array.isArray(item?.parsed?.visible_documents)
        ? item.parsed.visible_documents.map((doc) => ({
          ...doc,
          raw_document_label: String(doc?.document_label || "").trim(),
          document_label: canonicalLabel || String(doc?.document_label || "").trim(),
        }))
        : [];
    }),
    manual_checks: results.flatMap((item) => Array.isArray(item?.parsed?.manual_checks) ? item.parsed.manual_checks : []),
    notes: results.flatMap((item) => Array.isArray(item?.parsed?.notes) ? item.parsed.notes : []),
  };
  const tempPhytoContexts = [];
  for (const item of results) {
    const docs = Array.isArray(item?.job?.payload_json?.deterministic_visual_context?.temp_phyto_documents)
      ? item.job.payload_json.deterministic_visual_context.temp_phyto_documents
      : [];
    tempPhytoContexts.push(...docs);
  }
  const tempPhytoExpectedCount = tempPhytoContexts.length;
  const visiblePcnuPassChecks = combinedParsed.checks.filter((item) => item?.code === "PHYTO_PCNU_VISIBLE" && item?.status === "pass");
  const visiblePcnuDocumentLabels = new Set();
  for (const item of combinedParsed.visible_documents) {
    const pcnuValue = String(item?.pcnu_number || "").trim();
    const rawLabel = String(item?.document_label || "").trim();
    if (!pcnuValue || !rawLabel) {
      continue;
    }
    for (const label of rawLabel.split("|")) {
      const normalizedLabel = String(label || "").trim();
      if (normalizedLabel) {
        visiblePcnuDocumentLabels.add(normalizedLabel);
      }
    }
  }
  const combinedPcnuPassForAllDocs = visiblePcnuPassChecks.some((item) => {
    const normalizedMessage = String(item?.message || "").trim().toLowerCase();
    return normalizedMessage.includes("both temporary phyto documents")
      || normalizedMessage.includes("all temporary phyto documents")
      || normalizedMessage.includes("all three temporary phytosanitary certificates");
  });
  const visualPcnuCoveredAllTempPhytos = tempPhytoExpectedCount > 0 && (
    visiblePcnuPassChecks.length >= tempPhytoExpectedCount
    || visiblePcnuDocumentLabels.size >= tempPhytoExpectedCount
    || combinedPcnuPassForAllDocs
  );
  const deterministicChecks = (Array.isArray(deterministicSource.checks) ? deterministicSource.checks : []).map((item) => {
    if (
      item?.code === "TEMP_PHYTO_PARSE"
      && visualPcnuCoveredAllTempPhytos
      && String(item?.message || "").includes("missing a parsed PCNU number from PDF text extraction")
    ) {
      return {
        ...item,
        status: "pass",
        message: "PDF text extraction missed one or more PCNU numbers, but visual CSI confirmed all visible PCNU numbers.",
      };
    }
    return item;
  });
  const visualContextByLabel = new Map(
    tempPhytoContexts.map((item) => [normalizeUkdocsCsiDocumentLabel(item?.document_label), item]),
  );
  const visualTempPhytoTotals = new Map();
  const visualTempPhytoByLabel = new Map();
  for (const item of combinedParsed.visible_documents) {
    const documentLabel = String(item?.document_label || "").trim();
    const quantity = Number(item?.quantity);
    const noteText = String(item?.note || "").trim().toLowerCase();
    if (!documentLabel || !Number.isFinite(quantity)) {
      continue;
    }
    const context = visualContextByLabel.get(normalizeUkdocsCsiDocumentLabel(documentLabel));
    const mappedProduct = mapUkdocsCsiProductName(item?.product || "", "", {
      document_name: context?.name || "",
      prefer_plants: context?.prefer_plants === true,
      strict_domain: context?.prefer_plants === true ? "plants" : "flowers",
    });
    const parsedLineProducts = Array.isArray(context?.expected_products) ? context.expected_products : [];
    const needsFallback = !parsedLineProducts.length && context?.prefer_plants !== true;
    const mentionsTotal = noteText.includes("visible total") || noteText.includes("page total") || noteText.includes("total");
    if (!needsFallback || mentionsTotal || isUkdocsCsiAggregateProductName(mappedProduct)) {
      continue;
    }
    addUkdocsCsiQuantity(visualTempPhytoTotals, mappedProduct, quantity);
    if (!visualTempPhytoByLabel.has(documentLabel)) {
      visualTempPhytoByLabel.set(documentLabel, new Map());
    }
    const currentByProduct = visualTempPhytoByLabel.get(documentLabel).get(mappedProduct) || 0;
    visualTempPhytoByLabel.get(documentLabel).set(mappedProduct, currentByProduct + quantity);
  }
  const finalChecks = [
    ...deterministicChecks,
    ...combinedParsed.checks.filter((item) => item?.code !== "PHYTO_VISIBLE_QTY"),
  ];
  const deterministicProducts = Array.isArray(deterministicSource.products) ? deterministicSource.products : [];
  const finalProducts = deterministicProducts.map((item) => {
    const productName = String(item?.product || "").trim();
    const visualQty = productName && visualTempPhytoTotals.has(productName)
      ? visualTempPhytoTotals.get(productName)
      : null;
    const currentQty = String(item?.temp_phyto_quantity || "").trim();
    const currentPerDoc = Array.isArray(item?.temp_phyto_quantities) ? item.temp_phyto_quantities : [];
    const mergedPerDoc = currentPerDoc.map((entry) => {
      const label = String(entry?.document_label || "").trim();
      const qty = String(entry?.quantity || "").trim();
      if (qty || !label || !productName || !visualTempPhytoByLabel.has(label)) {
        return entry;
      }
      const visualDocQty = visualTempPhytoByLabel.get(label).get(productName);
      return visualDocQty === undefined
        ? entry
        : { ...entry, quantity: String(visualDocQty) };
    });
    for (const [label, productMap] of visualTempPhytoByLabel.entries()) {
      if (!productName || !productMap.has(productName) || mergedPerDoc.some((entry) => String(entry?.document_label || "").trim() === label)) {
        continue;
      }
      mergedPerDoc.push({
        document_label: label,
        quantity: String(productMap.get(productName)),
      });
    }
    if (!currentQty && visualQty !== null) {
      const existingMessage = String(item?.message || "").trim();
      return {
        ...item,
        temp_phyto_quantity: String(visualQty),
        temp_phyto_quantities: mergedPerDoc,
        message: existingMessage
          ? `${existingMessage} Visual fallback temp phyto quantity ${visualQty}.`
          : `Visual fallback temp phyto quantity ${visualQty}.`,
      };
    }
    return {
      ...item,
      temp_phyto_quantities: mergedPerDoc,
    };
  });
  const visualRows = combinedParsed.visible_documents.flatMap((item) => {
    const documentLabel = String(item?.document_label || "").trim();
    const context = visualContextByLabel.get(normalizeUkdocsCsiDocumentLabel(documentLabel));
    const hasParsedPlantRows = Array.isArray(deterministicSource.source_rows) && deterministicSource.source_rows.some((row) => (
      String(row?.source || "").trim() === "temp_phyto_plants"
      && normalizeUkdocsCsiDocumentLabel(row?.document_label || "") === normalizeUkdocsCsiDocumentLabel(documentLabel)
    ));
    if (context?.prefer_plants === true && hasParsedPlantRows) {
      return [];
    }
    const mappedProduct = mapUkdocsCsiProductName(item?.product || "", "", {
      document_name: context?.name || "",
      prefer_plants: context?.prefer_plants === true,
      strict_domain: context?.prefer_plants === true ? "plants" : "flowers",
    });
    return [{
      source: context?.prefer_plants === true ? "visual_temp_phyto_plants" : "visual_temp_phyto",
      document_name: String(context?.name || documentLabel).trim(),
      document_label: documentLabel,
      raw_product: String(item?.product || "").trim(),
      commodity_code: "",
      mapped_product: mappedProduct,
      product_domain: getUkdocsCsiProductDomain(mappedProduct),
      quantity: Number.isFinite(Number(item?.quantity)) ? Number(item.quantity) : null,
    }];
  });
  const visualDocumentLabels = new Set(
    visualRows.map((item) => normalizeUkdocsCsiDocumentLabel(item?.document_label || "")).filter(Boolean),
  );
  const finalSourceRows = [
    ...(Array.isArray(deterministicSource.source_rows) ? deterministicSource.source_rows : []).filter((row) => {
      const source = String(row?.source || "").trim();
      const isTempPhytoRow = source === "temp_phyto" || source === "temp_phyto_plants";
      if (!isTempPhytoRow) {
        return true;
      }
      const documentLabel = normalizeUkdocsCsiDocumentLabel(row?.document_label || "");
      if (!documentLabel || !visualDocumentLabels.has(documentLabel)) {
        return true;
      }
      if (source === "temp_phyto_plants") {
        return true;
      }
      return false;
    }),
    ...visualRows,
  ];
  const finalFlowerProducts = buildUkdocsCsiDomainProducts(finalProducts, finalSourceRows, "flowers");
  const finalPlantProducts = buildUkdocsCsiDomainProducts(finalProducts, finalSourceRows, "plants");
  const mergedProducts = [...finalFlowerProducts, ...finalPlantProducts].sort((left, right) => (
    String(left?.product || "").localeCompare(String(right?.product || ""))
  ));
  const finalManualChecks = uniqueUkdocsCsiStrings([
    ...(Array.isArray(deterministicSource.manual_checks) ? deterministicSource.manual_checks : []),
    ...combinedParsed.manual_checks,
  ]).filter((item) => !(visualPcnuCoveredAllTempPhytos && String(item || "").toLowerCase().includes("pcnu")));
  const finalNotes = uniqueUkdocsCsiStrings([
    ...(Array.isArray(deterministicSource.notes) ? deterministicSource.notes : []),
    ...combinedParsed.notes,
    visualPcnuCoveredAllTempPhytos
      ? "Visual CSI confirmed the visible PCNU number on every temporary phyto PDF, even where PDF text extraction missed it."
      : "",
  ]);
  const overallStatus = finalizeUkdocsCsiOverallStatus(finalChecks, mergedProducts);
  const deterministicSummary = String(deterministicSource.summary || "").trim();
  const finalSummary = uniqueUkdocsCsiStrings([
    deterministicSummary,
    combinedParsed.summary ? `Visual phyto check: ${combinedParsed.summary}` : "",
    ...results.map((item) => item.parseError ? `Visual phyto check incomplete: ${item.parseError}` : ""),
  ]).join(" ");
  return {
    status: "done",
    error: "",
    summary: finalSummary || deterministicSummary || "CSI audit completed.",
    overall_status: overallStatus,
    checks: finalChecks,
    products: mergedProducts,
    flower_products: finalFlowerProducts,
    plant_products: finalPlantProducts,
    source_rows: finalSourceRows,
    manual_checks: finalManualChecks,
    notes: finalNotes,
    llm_content: results.map((item) => {
      const label = String(item?.job?.payload_json?.csi_document_label || "").trim();
      const text = String(item?.llmContent || "").trim();
      return text ? `${label || item?.job?.id || "CSI"}\n${text}` : "";
    }).filter(Boolean).join("\n\n"),
    llm_parse_source: uniqueUkdocsCsiStrings(results.map((item) => item.parseSource || "")).join(", "),
    llm_parse_error: uniqueUkdocsCsiStrings(results.map((item) => item.parseError || "")).join(" | "),
    llm_raw_result_json: JSON.stringify(results.map((item) => ({
      job_id: item?.job?.id || "",
      document_label: item?.job?.payload_json?.csi_document_label || "",
      result_json: item?.job?.result_json || {},
    })), null, 2),
  };
}

async function queueUkdocsCsiAudit(collection, requestUser) {
  if (!isDatabaseEnabled()) {
    throw new Error("Database is not enabled");
  }
  if (!llmPollerEnabled()) {
    throw new Error("LLM poller is not configured");
  }

  const extractedDocuments = await extractUkdocsCsiFileSnapshots([
    ...(collection?.documents?.generated_files || []).map((document) => ({ kind: document.document_kind === "export" ? "generated_export" : "generated_invoice", document })),
    ...getUkdocsCsiTempPhytoParsedSourceDocuments(collection).map((item) => ({ kind: item.kind, document: item.document })),
    collection?.documents?.ipaffs_file ? [{ kind: "ipaffs_file", document: collection.documents.ipaffs_file }] : [],
    collection?.documents?.ipaffs_plants_file ? [{ kind: "ipaffs_plants_file", document: collection.documents.ipaffs_plants_file }] : [],
  ]);
  const deterministicBundle = buildUkdocsCsiDeterministicReport(collection, extractedDocuments);
  const deterministicVisualContextByLabel = new Map(
    (deterministicBundle?.visual_context?.temp_phyto_documents || []).map((item) => [String(item?.document_label || "").trim(), item]),
  );

  const tempPhytoVisionDocuments = await Promise.all(
    getUkdocsCsiTempPhytoVisionSourceDocuments(collection).map(async (item, index) => {
      const documentLabel = `Temp phyto ${String.fromCharCode(65 + index)}`;
      const deterministicContext = deterministicVisualContextByLabel.get(documentLabel) || null;
      if (deterministicContext?.deterministic_ready) {
        return null;
      }
      const document = item.document;
      const resolvedPath = path.resolve(ukdocsPrintDocumentPath(document));
      if (!resolvedPath.startsWith(path.resolve(ukdocsPrintFilesDir)) || !existsSync(resolvedPath)) {
        return null;
      }
      const contentBase64 = await fs.readFile(resolvedPath, "base64");
      return {
        document_label: documentLabel,
        name: String(document.original_name || document.storage_name || "temp-phyto.pdf").trim(),
        prefer_plants: item.prefer_plants === true,
        mime_type: String(document.mime_type || "application/pdf").trim(),
        content_base64: contentBase64,
        max_pages: 2,
      };
    }),
  );
  const usableVisionDocuments = tempPhytoVisionDocuments.filter(Boolean);

  if (!usableVisionDocuments.length) {
    await updateUkdocsCsiReport(collection.id, {
      status: "done",
      job_id: "",
      queued_at: new Date().toISOString(),
      started_at: new Date().toISOString(),
      completed_at: new Date().toISOString(),
      error: "",
      summary: deterministicBundle.report.summary || "CSI audit completed.",
      overall_status: deterministicBundle.report.overall_status || "warn",
      checks: deterministicBundle.report.checks || [],
      products: deterministicBundle.report.products || [],
      flower_products: deterministicBundle.report.flower_products || [],
      plant_products: deterministicBundle.report.plant_products || [],
      manual_checks: deterministicBundle.report.manual_checks || [],
      notes: uniqueUkdocsCsiStrings([
        ...(deterministicBundle.report.notes || []),
        (deterministicBundle?.visual_context?.temp_phyto_documents || []).some((item) => item?.deterministic_ready)
          ? "No temp phyto vision job was queued because deterministic temp phyto parsing was already sufficient for the available documents."
          : "No temp phyto vision job was queued because no readable temp phyto PDF was available on disk.",
      ]),
      llm_content: "",
      llm_parse_source: "",
      llm_parse_error: "",
      llm_raw_result_json: "",
    });
    await updateUkdocsCsiEmailResult(collection.id, {
      ok: false,
      recipients: [],
      sent_at: "",
      error: "",
    });
    return [];
  }

  const csiGroupId = crypto.randomUUID();
  const jobs = [];
  for (const visionDocument of usableVisionDocuments) {
    const visualContext = {
      temp_phyto_documents: (deterministicBundle?.visual_context?.temp_phyto_documents || [])
        .filter((item) => String(item?.document_label || "").trim() === String(visionDocument.document_label || "").trim()),
    };
    const job = await createLlmJob({
      job_type: "ukdocs_csi_audit",
      created_by: requestUser.username,
      collection_id: collection.id,
      shipment_id: collection.shipment_id,
      document_kind: "ukdocs_csi_audit",
      priority: 50,
      max_attempts: 1,
      payload_json: {
        ...buildUkdocsCsiAuditPayload(collection, deterministicBundle, requestUser, {
          vision_documents: [visionDocument],
          visual_context: visualContext,
        }),
        csi_group_id: csiGroupId,
        csi_document_label: visionDocument.document_label,
        csi_job_mode: "temp_phyto_single",
        deterministic_report: deterministicBundle.report,
        deterministic_visual_context: visualContext,
        vision_documents: [visionDocument],
      },
    });
    jobs.push(job);
  }

  await updateUkdocsCsiReport(collection.id, {
    status: "queued",
    job_id: jobs[0]?.id || "",
    queued_at: new Date().toISOString(),
    started_at: "",
    completed_at: "",
    error: "",
    summary: deterministicBundle.report.summary || "CSI audit queued.",
    overall_status: deterministicBundle.report.overall_status || "",
    checks: deterministicBundle.report.checks || [],
    products: deterministicBundle.report.products || [],
    flower_products: deterministicBundle.report.flower_products || [],
    plant_products: deterministicBundle.report.plant_products || [],
    manual_checks: deterministicBundle.report.manual_checks || [],
    notes: uniqueUkdocsCsiStrings([
      ...(deterministicBundle.report.notes || []),
      `Queued ${jobs.length} separate temp phyto CSI job(s).`,
    ]),
    llm_content: "",
    llm_parse_source: "",
    llm_parse_error: "",
    llm_raw_result_json: "",
  });
  await updateUkdocsCsiEmailResult(collection.id, {
    ok: false,
    recipients: [],
    sent_at: "",
    error: "",
  });

  return jobs;
}

function ukdocsPrintCollectionMatchScore(collection, haystackRaw) {
  const haystack = normalizeUkdocsPrintToken(haystackRaw);
  if (!haystack) {
    return 0;
  }
  let score = 0;
  const referenceConnects = ukdocsPrintReferenceTokens(collection?.reference_connect);
  for (const referenceConnect of referenceConnects) {
    if (referenceConnect && haystack.includes(referenceConnect)) {
      score += 8;
    }
  }
  const invoices = ukdocsPrintInvoiceTokens(collection?.invoice_numbers);
  for (const invoice of invoices) {
    if (invoice && haystack.includes(invoice)) {
      score += 5;
    }
  }
  const truck = normalizeUkdocsPrintToken(collection?.truck_number);
  const trailer = normalizeUkdocsPrintToken(collection?.trailer_number);
  if (truck && haystack.includes(truck)) {
    score += 2;
  }
  if (trailer && haystack.includes(trailer)) {
    score += 2;
  }
  return score;
}

function collectionAcceptsUkdocsPrintDocument(collection, customers, kind) {
  const inspectionMode = ukdocsPrintInspectionMode(collection);
  const customer = ukdocsPrintCollectionCustomer(collection, customers);
  const pdTypeCompact = String(collection?.pd_type || "").trim().toLowerCase().replace(/\s+/g, "");

  if (kind === "phyto") {
    if (inspectionMode === "stock_control") {
      return false;
    }
    if (customer?.required_phyto === false) {
      return false;
    }
    if (pdTypeCompact.includes("nophytoneeded")) {
      return false;
    }
    return true;
  }

  return inspectionMode !== "stock_control";
}

function findUkdocsPrintDocumentOwner(collections, date, kind, originalName) {
  const syncDate = String(date || "").slice(0, 10);
  const target = String(originalName || "").trim().toLowerCase();
  if (!syncDate || !target) {
    return null;
  }
  for (const collection of Array.isArray(collections) ? collections : []) {
    if (String(collection?.shipment_date || "").slice(0, 10) !== syncDate) {
      continue;
    }
    if (kind === "phyto") {
      const files = normalizeUkdocsPrintDocumentList(collection?.documents?.phyto_files);
      const index = files.findIndex((document) => ukdocsPrintDocumentIdentity(document) === target);
      if (index >= 0) {
        return { collection, document: files[index], index };
      }
      continue;
    }
    const document = normalizeUkdocsPrintDocument(collection?.documents?.[kind]);
    if (document && ukdocsPrintDocumentIdentity(document) === target) {
      return { collection, document, index: 0 };
    }
  }
  return null;
}

function removeUkdocsPrintDocumentByName(collection, kind, originalName) {
  const target = String(originalName || "").trim().toLowerCase();
  if (!target) {
    return normalizeUkdocsPrintCollection(collection);
  }
  if (kind === "phyto") {
    return normalizeUkdocsPrintCollection({
      ...collection,
      updated_at: new Date().toISOString(),
      documents: {
        ...(collection?.documents || {}),
        phyto_files: normalizeUkdocsPrintDocumentList(collection?.documents?.phyto_files).filter((document) => ukdocsPrintDocumentIdentity(document) !== target),
      },
    });
  }
  const currentDocument = normalizeUkdocsPrintDocument(collection?.documents?.[kind]);
  if (currentDocument && ukdocsPrintDocumentIdentity(currentDocument) === target) {
    return normalizeUkdocsPrintCollection({
      ...collection,
      updated_at: new Date().toISOString(),
      documents: {
        ...(collection?.documents || {}),
        [kind]: null,
      },
    });
  }
  return normalizeUkdocsPrintCollection(collection);
}

function removeUkdocsPrintDocumentFromOtherCollections(collections, date, targetCollectionId, kind, originalName) {
  const syncDate = String(date || "").slice(0, 10);
  const target = String(originalName || "").trim().toLowerCase();
  if (!syncDate || !target) {
    return Array.isArray(collections) ? collections : [];
  }

  return (Array.isArray(collections) ? collections : []).map((collection) => {
    if (String(collection?.shipment_date || "").slice(0, 10) !== syncDate) {
      return collection;
    }
    if (collection?.id === targetCollectionId) {
      return collection;
    }
    return removeUkdocsPrintDocumentByName(collection, kind, originalName);
  });
}

function detectUkdocsPrintDocumentKind(text) {
  const normalized = String(text || "").toLowerCase();
  if (/(phyto|phytosan|kcb|certificate|certificaat|e-certnl|nvwa\.nl|no-reply@nvwa\.nl)/.test(normalized)) {
    return "phyto";
  }
  return "export_extra";
}

async function gmailApiJson(accessToken, resourcePath) {
  const response = await fetch(`https://gmail.googleapis.com/gmail/v1/users/me/${resourcePath}`, {
    headers: { authorization: `Bearer ${accessToken}` },
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(payload.error?.message || payload.error_description || `Gmail API failed with ${response.status}`);
  }
  return payload;
}

async function gmailAttachmentBuffer(accessToken, messageId, attachmentId) {
  const payload = await gmailApiJson(accessToken, `messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}`);
  const base64Url = String(payload.data || "").trim();
  if (!base64Url) {
    return Buffer.alloc(0);
  }
  return Buffer.from(base64Url.replace(/-/g, "+").replace(/_/g, "/"), "base64");
}

function gmailHeaderValue(headers, name) {
  const header = Array.isArray(headers) ? headers.find((item) => String(item?.name || "").toLowerCase() === String(name || "").toLowerCase()) : null;
  return String(header?.value || "");
}

function collectGmailAttachments(parts, bucket = []) {
  for (const part of Array.isArray(parts) ? parts : []) {
    if (part?.body?.attachmentId && part?.filename) {
      bucket.push({
        filename: String(part.filename || ""),
        mime_type: String(part.mimeType || ""),
        attachment_id: String(part.body.attachmentId || ""),
      });
    }
    if (Array.isArray(part?.parts) && part.parts.length) {
      collectGmailAttachments(part.parts, bucket);
    }
  }
  return bucket;
}

function formatGmailQueryDate(date) {
  const value = String(date || "").slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    return "";
  }
  return value.replace(/-/g, "/");
}

function nextIsoDate(date) {
  const value = String(date || "").slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    return "";
  }
  const current = new Date(`${value}T12:00:00`);
  current.setDate(current.getDate() + 1);
  return current.toISOString().slice(0, 10);
}

function buildUkdocsGmailSyncQuery(baseQuery, date) {
  const trimmedQuery = String(baseQuery || "").trim() || "has:attachment";
  const startDate = formatGmailQueryDate(date);
  const nextDate = formatGmailQueryDate(nextIsoDate(date));
  if (!startDate || !nextDate) {
    return trimmedQuery;
  }
  return `${trimmedQuery} after:${startDate} before:${nextDate}`.trim();
}

async function syncUkdocsPrintFromGmail(settings, requestUser, query, date) {
  if (!settings.gmail_refresh_token) {
    throw new Error("Connect Gmail first");
  }
  const accessToken = await refreshGoogleAccessToken(settings, settings.gmail_refresh_token);
  const state = await readUkdocsState();
  const syncDate = String(date || localDateIso()).slice(0, 10);
  const syncQuery = buildUkdocsGmailSyncQuery(query, syncDate);
  const listPayload = await gmailApiJson(accessToken, `messages?q=${encodeURIComponent(syncQuery)}&maxResults=25`);
  const messages = Array.isArray(listPayload.messages) ? listPayload.messages : [];
  const dayCollections = state.print_collections.filter((item) => {
    if (String(item.shipment_date || "").slice(0, 10) !== syncDate) {
      return false;
    }
    const inspectionMode = ukdocsPrintInspectionMode(item);
    return inspectionMode !== "stock_control";
  });
  const results = [];

  for (const message of messages) {
    const detail = await gmailApiJson(accessToken, `messages/${encodeURIComponent(message.id)}?format=full`);
    const subject = gmailHeaderValue(detail.payload?.headers, "subject");
    const fromHeader = gmailHeaderValue(detail.payload?.headers, "from");
    const textBlob = [subject, fromHeader, detail.snippet, detail.id].join(" ");
    const attachments = collectGmailAttachments(detail.payload?.parts || []);
    for (const attachment of attachments) {
      const attachmentName = attachment.filename || "attachment";
      if (!/\.(pdf|xls|xlsx)$/i.test(attachmentName)) {
        results.push({ status: "skipped", file_name: attachmentName, reason: "Unsupported file type" });
        continue;
      }
      const candidateText = `${textBlob} ${attachmentName}`;
      const kind = detectUkdocsPrintDocumentKind(candidateText);
      const eligibleCollections = dayCollections.filter((collection) => collectionAcceptsUkdocsPrintDocument(collection, state.customers, kind));
      const ranked = eligibleCollections
        .map((collection) => ({ collection, score: ukdocsPrintCollectionMatchScore(collection, candidateText) }))
        .filter((item) => item.score > 0)
        .sort((a, b) => b.score - a.score);
      const bestScore = ranked[0]?.score || 0;
      const bestMatch = bestScore >= 5 && (ranked.length === 1 || ranked[0].score > ranked[1].score)
        ? ranked[0].collection
        : null;
      if (!bestMatch) {
        results.push({ status: "unmatched", file_name: attachmentName, reason: "No safe match from reference connect, invoice, or truck/trailer" });
        continue;
      }
      const currentMatch = state.print_collections.find((item) => item.id === bestMatch.id || item.shipment_id === bestMatch.shipment_id) || bestMatch;
      const existingOwner = findUkdocsPrintDocumentOwner(state.print_collections, syncDate, kind, attachmentName);
      state.print_collections = removeUkdocsPrintDocumentFromOtherCollections(state.print_collections, syncDate, currentMatch.id, kind, attachmentName);
      if (existingOwner?.collection?.id === currentMatch.id) {
        const refreshedCurrent = state.print_collections.find((item) => item.id === currentMatch.id) || currentMatch;
        state.print_collections = upsertUkdocsPrintCollection(state.print_collections, refreshedCurrent);
        results.push({ status: "skipped", file_name: attachmentName, shipment_reference: currentMatch.shipment_reference, reason: `${kind} already exists` });
        continue;
      }
      if (kind !== "phyto" && currentMatch.documents?.[kind]?.storage_name && !existingOwner) {
        results.push({ status: "skipped", file_name: attachmentName, shipment_reference: currentMatch.shipment_reference, reason: `${kind} already exists` });
        continue;
      }
      let savedDocument = existingOwner?.document || null;
      if (!savedDocument) {
        const buffer = await gmailAttachmentBuffer(accessToken, detail.id, attachment.attachment_id);
        savedDocument = await saveUkdocsPrintBuffer(currentMatch.id, kind, attachmentName, attachment.mime_type, buffer, requestUser.username);
      }
      const updatedCollection = normalizeUkdocsPrintCollection({
        ...currentMatch,
        updated_at: new Date().toISOString(),
        documents: {
          ...(currentMatch.documents || {}),
          ...(kind === "phyto"
            ? { phyto_files: [...(currentMatch.documents?.phyto_files || []), savedDocument] }
            : { [kind]: savedDocument }),
        },
      });
      state.print_collections = upsertUkdocsPrintCollection(state.print_collections, updatedCollection);
      results.push({
        status: "matched",
        file_name: attachmentName,
        shipment_reference: updatedCollection.shipment_reference,
        kind,
        reason: existingOwner && existingOwner.collection.id !== currentMatch.id ? "moved to better match" : "-",
      });
    }
  }

  await writeUkdocsState(state);
  return {
    date: syncDate,
    query: syncQuery,
    matched: results.filter((item) => item.status === "matched").length,
    unmatched: results.filter((item) => item.status === "unmatched").length,
    skipped: results.filter((item) => item.status === "skipped").length,
    results,
    print_collections: normalizeUkdocsState(state).print_collections,
  };
}

async function syncUkdocsPrintCollectionsFromSheet(settings, date, options = {}) {
  const spreadsheetId = String(settings.ukdocs_print_spreadsheet_id || "").trim();
  const sheetName = String(settings.ukdocs_print_sheet_name || defaultFustSettings.ukdocs_print_sheet_name).trim();
  if (!spreadsheetId || !sheetName) {
    throw new Error("Set a UKdocs Print spreadsheet ID and tab name in Settings first");
  }
  const rows = await loadSheetRows(spreadsheetId, sheetName);
  const sendings = parseUkdocsPrintSheetRows(rows, date);
  const state = await readUkdocsState();
  const referenceConnectOnly = options.reference_connect_only === true;
  const updateOnly = options.update_only === true;
  const overwriteExisting = options.overwrite_existing === true;
  const syncDate = String(date || localDateIso()).slice(0, 10);
  if (overwriteExisting) {
    state.print_collections = (Array.isArray(state.print_collections) ? state.print_collections : []).filter((item) => {
      if (String(item?.shipment_date || "").slice(0, 10) !== syncDate) {
        return true;
      }
      return item?.collection_type === "stock_control";
    });
  }
  let updatedCount = 0;
  for (const sending of sendings) {
    const existingCollection = findMatchingUkdocsPrintCollection(state.print_collections, sending, { allowInvoiceFallback: false });
    const matchedCustomer = matchUkdocsCustomerForPrintCollection(state.customers, sending);
    const nextCollection = referenceConnectOnly && existingCollection
      ? normalizeUkdocsPrintCollection({
        ...existingCollection,
        reference_connect: sending.reference_connect || existingCollection.reference_connect || "",
        sheet_row_number: sending.sheet_row_number || existingCollection.sheet_row_number || 0,
        updated_at: new Date().toISOString(),
      })
      : updateOnly
        ? normalizeUkdocsPrintCollection({
          ...existingCollection,
          id: existingCollection?.id || sending.id,
          source: "sheet",
          shipment_date: sending.shipment_date,
          customer_id: existingCollection?.customer_id || matchedCustomer?.id || "",
          customer_name: matchedCustomer?.customer_name || existingCollection?.customer_name || sending.city_name || "",
          collection_type: existingCollection?.collection_type || sending.collection_type || (isHonselersdijkStockControl(sending) ? "stock_control" : "export"),
          city_name: sending.city_name,
          border_crossing: sending.border_crossing,
          hub_code: sending.hub_code,
          remark: sending.remark,
          pd_form: sending.pd_form,
          re_export: sending.re_export,
          pd_type: sending.pd_type,
          pd_code: sending.pd_code,
          reference_connect: sending.reference_connect,
          trailer_number: sending.trailer_number || "",
          truck_number: sending.truck_number || "",
          invoice_numbers: sending.invoice_numbers || "",
          sheet_row_number: sending.sheet_row_number,
          updated_at: new Date().toISOString(),
          generated_at: existingCollection?.generated_at || "",
          documents: existingCollection?.documents || {},
          notes: existingCollection?.notes || "",
          delivery_email: existingCollection?.delivery_email || {},
        })
        : normalizeUkdocsPrintCollection({
          ...existingCollection,
          id: existingCollection?.id || sending.id,
          source: "sheet",
          shipment_date: sending.shipment_date,
          customer_id: existingCollection?.customer_id || matchedCustomer?.id || "",
          customer_name: existingCollection?.customer_name || matchedCustomer?.customer_name || sending.city_name || "",
          collection_type: existingCollection?.collection_type || sending.collection_type || (isHonselersdijkStockControl(sending) ? "stock_control" : "export"),
          city_name: sending.city_name,
          border_crossing: sending.border_crossing,
          hub_code: sending.hub_code,
          remark: sending.remark,
          pd_form: sending.pd_form,
          re_export: sending.re_export,
          pd_type: sending.pd_type,
          pd_code: sending.pd_code,
          reference_connect: sending.reference_connect,
          trailer_number: existingCollection?.trailer_number || sending.trailer_number || "",
          truck_number: existingCollection?.truck_number || sending.truck_number || "",
          invoice_numbers: existingCollection?.invoice_numbers || sending.invoice_numbers || "",
          sheet_row_number: sending.sheet_row_number,
          updated_at: new Date().toISOString(),
          generated_at: existingCollection?.generated_at || "",
          documents: existingCollection?.documents || {},
          notes: existingCollection?.notes || "",
        });
    state.print_collections = upsertUkdocsPrintCollection(state.print_collections, nextCollection);
    updatedCount += 1;
  }
  dedupeUkdocsPrintCollectionsForDate(state, syncDate);
  await writeUkdocsState(state);
  return {
    spreadsheet_id: spreadsheetId,
    sheet_name: sheetName,
    date: syncDate,
    imported_count: sendings.length,
    updated_count: updatedCount,
    reference_connect_only: referenceConnectOnly,
    update_only: updateOnly,
    overwrite_existing: overwriteExisting,
    print_collections: normalizeUkdocsState(state).print_collections,
  };
}


function runHalLocationsWorker(args) {
  return new Promise((resolve, reject) => {
    const child = spawn(resolvePythonCommand(), [halLocationsWorkerPath, ...args], {
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

      reject(new Error(summarizeBridgeError(Buffer.concat(stderr).toString("utf8")) || `Hal Locations worker exited with ${code}`));
    });

    child.stdin.end();
  });
}

function runExpeditionStickerWorker(args) {
  return new Promise((resolve, reject) => {
    const child = spawn(resolvePythonCommand(), [expeditionStickerWorkerPath, ...args], {
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

      reject(new Error(summarizeBridgeError(Buffer.concat(stderr).toString("utf8")) || `Expedition Sticker worker exited with ${code}`));
    });

    child.stdin.end();
  });
}

function runFustImportWorker(args) {
  return new Promise((resolve, reject) => {
    const child = spawn(resolvePythonCommand(), [fustImportWorkerPath, ...args], {
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

      reject(new Error(summarizeBridgeError(Buffer.concat(stderr).toString("utf8")) || `Fust import worker exited with ${code}`));
    });

    child.stdin.end();
  });
}

function runFustListWorker(args) {
  return new Promise((resolve, reject) => {
    const child = spawn(resolvePythonCommand(), [fustListWorkerPath, ...args], {
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

      reject(new Error(summarizeBridgeError(Buffer.concat(stderr).toString("utf8")) || `Fust list worker exited with ${code}`));
    });

    child.stdin.end();
  });
}

function findFustListTemplatePath() {
  return fustListTemplatePathCandidates.find((candidate) => existsSync(candidate)) || "";
}

function normalizeFustListRow(row) {
  const code = String(row?.code || "").trim().toUpperCase();
  const totalOk = Math.max(0, Number.parseInt(String(row?.total_ok ?? 0), 10) || 0);
  const totalBroken = Math.max(0, Number.parseInt(String(row?.total_broken ?? 0), 10) || 0);
  return {
    code,
    total_ok: totalOk,
    total_broken: totalBroken,
  };
}

async function generateFustListWorkbook(payload) {
  const templatePath = findFustListTemplatePath();
  if (!templatePath) {
    throw new Error("Fust Lijst template file is missing");
  }

  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "fust-list-"));
  const payloadPath = path.join(tempDir, "payload.json");
  const outputPath = path.join(tempDir, "fust-lijst.xlsx");

  try {
    await fs.writeFile(payloadPath, JSON.stringify(payload), "utf8");
    await runFustListWorker([
      "generate",
      "--template",
      templatePath,
      "--payload",
      payloadPath,
      "--output",
      outputPath,
    ]);
    return await fs.readFile(outputPath);
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true }).catch(() => {});
  }
}

function normalizeFustExporterInfo(info) {
  return {
    name: String(info?.name || "").trim(),
    block: String(info?.block || "").trim(),
  };
}

async function cleanupExpiredHalLocationSessions() {
  const now = Date.now();
  for (const [sessionId, session] of halLocationSessions.entries()) {
    if (now - Number(session.ts || 0) <= halLocationSessionTtlMs) {
      continue;
    }
    halLocationSessions.delete(sessionId);
    if (session.dir_path) {
      await fs.rm(session.dir_path, { recursive: true, force: true }).catch(() => {});
    }
  }
}

async function createHalLocationSession(filePayload) {
  await cleanupExpiredHalLocationSessions();

  const fileName = path.basename(String(filePayload?.name || "halindeling.xlsx")).replace(/[^a-zA-Z0-9._ -]+/g, "_") || "halindeling.xlsx";
  const contentBase64 = String(filePayload?.content_base64 || "").trim();
  if (!contentBase64) {
    throw new Error("Upload a halindeling file first");
  }

  const sessionId = crypto.randomUUID();
  const dirPath = path.join(halLocationsCacheDir, sessionId);
  const filePath = path.join(dirPath, fileName);
  await fs.mkdir(dirPath, { recursive: true });
  await fs.writeFile(filePath, Buffer.from(contentBase64, "base64"));

  halLocationSessions.set(sessionId, {
    id: sessionId,
    dir_path: dirPath,
    file_path: filePath,
    file_name: fileName,
    source: "upload",
    ts: Date.now(),
  });
  return halLocationSessions.get(sessionId);
}

async function createHalLocationSheetSession(rows, sourceMeta = {}) {
  await cleanupExpiredHalLocationSessions();

  if (!Array.isArray(rows) || !rows.length) {
    throw new Error("No rows found in the Hal Locations sheet");
  }

  const sessionId = crypto.randomUUID();
  const dirPath = path.join(halLocationsCacheDir, sessionId);
  const filePath = path.join(dirPath, "ERP_PASTE.json");
  await fs.mkdir(dirPath, { recursive: true });
  await fs.writeFile(filePath, JSON.stringify(rows), "utf8");

  halLocationSessions.set(sessionId, {
    id: sessionId,
    dir_path: dirPath,
    file_path: filePath,
    file_name: "ERP_PASTE.json",
    source: "sheet",
    source_meta: sourceMeta,
    ts: Date.now(),
  });
  return halLocationSessions.get(sessionId);
}

function normalizeHalPrefixList(values) {
  if (!Array.isArray(values)) {
    return [];
  }
  return [...new Set(values.map((value) => String(value || "").trim()).filter(Boolean))];
}

async function getHalLocationSession(sessionId) {
  await cleanupExpiredHalLocationSessions();
  const session = halLocationSessions.get(String(sessionId || ""));
  if (!session) {
    return null;
  }
  session.ts = Date.now();
  return session;
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

function emptySheetRow(length) {
  return Array.from({ length: Math.max(1, Number(length) || 1) }, () => "");
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
  if (mimeType === "text/csv" || mimeType === "application/csv" || String(mimeType || "").toLowerCase().includes("csv")) {
    return ".csv";
  }
  if (mimeType === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") {
    return ".xlsx";
  }
  if (mimeType === "application/vnd.ms-excel") {
    return ".xls";
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

async function applyFustDocumentChoice(action, settings, documentPayload, requestUser) {
  const documentConfig = action.type === "IN"
    ? { field: "fustbon", label: "Fustbon" }
    : { field: "cmr", label: "CMR" };
  const mode = String(documentPayload?.mode || "").trim().toLowerCase();
  if (!mode) {
    return;
  }

  if (mode === "skip") {
    action[documentConfig.field] = normalizeCmrInfo({
      status: "skipped",
      uploaded_at: new Date().toISOString(),
      uploaded_by: requestUser.username,
    });
    return;
  }

  if (mode !== "upload") {
    throw new Error(`Unknown ${documentConfig.label} choice`);
  }

  const filePayload = documentPayload?.file || {};
  if (!filePayload.content_base64 || !filePayload.name) {
    throw new Error(`Choose a ${documentConfig.label} file first`);
  }

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
  }
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

async function clearFustActionRowInSheet(spreadsheetId, sheetName, type, action, explicitRowNumber = 0) {
  if (!spreadsheetId || !sheetName) {
    return 0;
  }

  const rows = await loadSheetRows(spreadsheetId, sheetName);
  const layout = getRegistrySheetLayout(rows);
  let targetRowNumber = Number(explicitRowNumber || action?.sheet_sync?.row_number || 0);
  if (targetRowNumber < 2 && action?.id) {
    targetRowNumber = findFustSheetRowNumberByActionId(rows, action.id);
  }
  if (targetRowNumber < 2) {
    targetRowNumber = findFustSheetRowNumberBySignature(rows, type, action);
  }
  if (targetRowNumber < 2) {
    return 0;
  }

  const existingRow = rows[targetRowNumber - 1] || [];
  const rowLength = Math.max(layout.rowLength, existingRow.length, 19);
  await writeSheetRowAt(spreadsheetId, sheetName, targetRowNumber, emptySheetRow(rowLength));
  return targetRowNumber;
}

async function deleteFustActionFromSheets(action, settings) {
  if (!settings.spreadsheet_id) {
    return { ok: false, target_sheets: [], error: "Spreadsheet ID is not configured" };
  }

  const clearTargets = [
    { type: "IN", sheetName: settings.in_sheet_name, explicitRowNumber: action?.type === "IN" ? action?.sheet_sync?.row_number : 0 },
    { type: "OUT", sheetName: settings.out_sheet_name, explicitRowNumber: action?.type === "OUT" ? action?.sheet_sync?.row_number : 0 },
  ].filter((item) => item.sheetName);

  const clearedSheets = [];
  for (const target of clearTargets) {
    const clearedRowNumber = await clearFustActionRowInSheet(
      settings.spreadsheet_id,
      target.sheetName,
      target.type,
      action,
      target.explicitRowNumber,
    );
    if (clearedRowNumber >= 2) {
      clearedSheets.push(target.sheetName);
    }
  }

  return {
    ok: true,
    target_sheets: clearedSheets,
    error: "",
    synced_at: new Date().toISOString(),
    row_number: 0,
  };
}

async function syncFustActionToSheets(action, settings, options = {}) {
  if (!settings.spreadsheet_id) {
    return { ok: false, target_sheets: [], error: "Spreadsheet ID is not configured" };
  }

  const targetSheet = action.type === "OUT" ? settings.out_sheet_name : settings.in_sheet_name;
  if (!targetSheet) {
    return { ok: false, target_sheets: [], error: "Target sheet is not configured" };
  }

  const previousAction = options.previousAction || null;
  const previousTargetSheet = previousAction
    ? (previousAction.type === "OUT" ? settings.out_sheet_name : settings.in_sheet_name)
    : "";
  if (previousAction && previousTargetSheet && previousTargetSheet !== targetSheet) {
    await clearFustActionRowInSheet(
      settings.spreadsheet_id,
      previousTargetSheet,
      previousAction.type,
      previousAction,
      previousAction?.sheet_sync?.row_number || 0,
    );
  }

  const existingRows = await loadSheetRows(settings.spreadsheet_id, targetSheet);
  let targetRowNumber = Number(action?.sheet_sync?.row_number || 0);
  if (targetRowNumber < 2) {
    if (action?.id) {
      targetRowNumber = findFustSheetRowNumberByActionId(existingRows, action.id);
    }
    if (targetRowNumber < 2) {
      targetRowNumber = findFustSheetRowNumberBySignature(existingRows, action.type, options.previousAction || action);
    }
  }

  const existingRow = targetRowNumber >= 1 ? (existingRows[targetRowNumber - 1] || []) : [];
  const rowPayload = buildRegistrySheetRow(action, existingRows, existingRow);
  const output = targetRowNumber >= 2
    ? await writeSheetRowAt(settings.spreadsheet_id, targetSheet, targetRowNumber, rowPayload)
    : await writeSheetRowToFirstEmpty(settings.spreadsheet_id, targetSheet, rowPayload);

  return {
    ok: true,
    target_sheets: [targetSheet],
    error: "",
    synced_at: new Date().toISOString(),
    row_number: Number(output?.row_number || targetRowNumber || 0),
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
    `Fustbon: ${action.fustbon_reference || "-"}`,
    `Fustfactuur: ${action.fustfactuur_reference || "-"}`,
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

function addDaysToIsoDate(dateString, days) {
  const parsed = new Date(`${String(dateString || "").slice(0, 10)}T00:00:00`);
  if (Number.isNaN(parsed.getTime())) {
    return "";
  }
  parsed.setDate(parsed.getDate() + Number(days || 0));
  return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, "0")}-${String(parsed.getDate()).padStart(2, "0")}`;
}

function fustActionControlDate(action) {
  const createdDate = String(action?.created_at || "").slice(0, 10);
  if (createdDate) {
    return createdDate;
  }
  const actionDate = String(action?.action_date || "").slice(0, 10);
  const importedDate = String(action?.import_source?.imported_at || "").slice(0, 10);
  if (importedDate && (!actionDate || importedDate > actionDate)) {
    return importedDate;
  }
  return actionDate;
}

function fustActionConfirmationDueDate(action) {
  return addDaysToIsoDate(fustActionControlDate(action), 1);
}

function isFustConfirmationReminderOverdue(action, todayIso = localDateIso()) {
  if (isFustActionConfirmed(action)) {
    return false;
  }
  const confirmByDate = fustActionConfirmationDueDate(action);
  if (!confirmByDate) {
    return false;
  }
  return String(todayIso || "") >= confirmByDate;
}

function buildFustConfirmationReminderEmail(actions) {
  const sortedActions = [...(Array.isArray(actions) ? actions : [])].sort((left, right) => {
    const leftDate = String(fustActionControlDate(left) || "");
    const rightDate = String(fustActionControlDate(right) || "");
    return leftDate.localeCompare(rightDate) || String(left.customer_name || "").localeCompare(String(right.customer_name || ""));
  });
  return [
    "Fust confirmation reminder",
    "",
    "These actions are still not confirmed and need attention:",
    "",
    ...sortedActions.map((action) => {
      const confirmByDate = fustActionConfirmationDueDate(action) || "-";
      const controlDate = fustActionControlDate(action) || action.action_date || "-";
      return [
        `${action.type} | ${action.action_date} | ${action.country} | ${action.customer_name}`,
        `Connect: ${action.connect_name || "-"}`,
        `Control date: ${controlDate}`,
        `Confirm by: ${confirmByDate}`,
        `DC: ${Number(action.metrics?.dc || 0)} | DCS: ${Number(action.metrics?.dcs || 0)} | DCO: ${Number(action.metrics?.dco || 0)} | CCTag: ${Number(action.metrics?.cctag || 0)} | PAL: ${Number(action.metrics?.pal || 0)} | VK: ${Number(action.metrics?.vk || 0)}`,
        action.remark ? `Remark: ${action.remark}` : "",
        "",
      ].filter(Boolean).join("\n");
    }),
  ].flat().join("\n");
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

async function sendFustConfirmationReminderEmail(actions, settings) {
  const recipients = normalizeEmailRecipients(settings.email_recipients);
  if (!recipients.length) {
    return { ok: false, recipients: [], error: "No email recipients configured" };
  }
  const subject = `Fust confirmation reminder | ${actions.length} overdue action${actions.length === 1 ? "" : "s"}`;
  await runPythonBridge(
    ["email-send"],
    JSON.stringify({
      recipients,
      subject,
      body: buildFustConfirmationReminderEmail(actions),
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

async function loadCurrentFustActionsSnapshot(settingsOverride = null) {
  const settings = settingsOverride || await readFustSettings();
  const localActions = await readFustActions();
  let inSheetActions = [];
  let outSheetActions = [];
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

  const deletedActionIds = new Set(
    localActions
      .filter((action) => action.deleted)
      .map((action) => String(action.id || "").trim())
      .filter(Boolean),
  );
  const dedupedActions = new Map();
  for (const action of [...inSheetActions, ...outSheetActions, ...localActions]) {
    const actionId = String(action?.id || "").trim();
    if (actionId && deletedActionIds.has(actionId)) {
      continue;
    }
    if (action.deleted) {
      continue;
    }
    dedupedActions.set(buildActionMergeKey(action), action);
  }

  const actions = [...dedupedActions.values()];
  return { settings, localActions, actions, sourceDebug };
}

async function maybeSendFustConfirmationReminders(actions, settings) {
  const today = localDateIso();
  const unconfirmedActions = (Array.isArray(actions) ? actions : []).filter((action) => !isFustActionConfirmed(action) && !isImportedFustAction(action));
  const yesterdayDate = addDaysToIsoDate(today, -1);
  const yesterdayUnconfirmedActions = unconfirmedActions.filter((action) => fustActionControlDate(action) === yesterdayDate);
  const overdueActions = unconfirmedActions.filter((action) => isFustConfirmationReminderOverdue(action, today));
  const reminderCandidates = overdueActions.filter((action) => String(action?.confirmation_reminder?.last_attempt_for_date || "") !== today);

  const summary = {
    checked_at: new Date().toISOString(),
    yesterday_date: yesterdayDate,
    yesterday_unconfirmed_count: yesterdayUnconfirmedActions.length,
    unconfirmed_count: unconfirmedActions.length,
    overdue_unconfirmed_count: overdueActions.length,
    reminder_candidate_count: reminderCandidates.length,
    reminder_sent: false,
    reminder_sent_at: "",
    reminder_error: "",
  };

  if (!reminderCandidates.length) {
    return summary;
  }

  const localActions = await readFustActions();
  const localById = new Map(localActions.map((action, index) => [String(action.id || "").trim(), { action, index }]).filter(([id]) => id));
  const updatedIndexes = new Set();

  function ensureReminderTarget(sourceAction) {
    const actionId = String(sourceAction?.id || "").trim();
    if (localById.has(actionId)) {
      return localById.get(actionId);
    }
    const seededAction = normalizeFustAction({
      ...sourceAction,
      created_by: sourceAction?.created_by === "spreadsheet" ? "sheet-import" : (sourceAction?.created_by || "sheet-import"),
      created_at: sourceAction?.created_at || new Date().toISOString(),
    });
    localActions.push(seededAction);
    const entry = { action: seededAction, index: localActions.length - 1 };
    localById.set(actionId, entry);
    return entry;
  }

  try {
    const emailResult = await sendFustConfirmationReminderEmail(reminderCandidates, settings);
    summary.reminder_sent = emailResult.ok === true;
    summary.reminder_sent_at = emailResult.sent_at || "";

    for (const action of reminderCandidates) {
      const target = ensureReminderTarget(action);
      target.action.confirmation_reminder = normalizeFustConfirmationReminder({
        ...(target.action.confirmation_reminder || {}),
        last_attempt_at: emailResult.sent_at || new Date().toISOString(),
        last_attempt_for_date: today,
        last_sent_at: emailResult.sent_at || new Date().toISOString(),
        last_sent_for_date: today,
        sent_count: Number(target.action?.confirmation_reminder?.sent_count || 0) + 1,
        last_error: "",
      });
      localActions[target.index] = normalizeFustAction(target.action);
      updatedIndexes.add(target.index);
    }
  } catch (error) {
    summary.reminder_error = error instanceof Error ? error.message : String(error || "Unknown reminder error");
    for (const action of reminderCandidates) {
      const target = ensureReminderTarget(action);
      target.action.confirmation_reminder = normalizeFustConfirmationReminder({
        ...(target.action.confirmation_reminder || {}),
        last_attempt_at: new Date().toISOString(),
        last_attempt_for_date: today,
        last_error: summary.reminder_error,
      });
      localActions[target.index] = normalizeFustAction(target.action);
      updatedIndexes.add(target.index);
    }
  }

  if (updatedIndexes.size) {
    await writeFustActions(localActions);
    for (const index of updatedIndexes) {
      await mirrorFustActionToDatabase(localActions[index]);
    }
  }

  return summary;
}

async function ukdocsPrintCollectionAttachments(collection, customer = null, menuKey = "ukdocsprint") {
  const attachments = [];
  const visibility = ukdocsMenuDocumentVisibility(customer, menuKey);
  const generatedFiles = Array.isArray(collection?.documents?.generated_files) ? collection.documents.generated_files : [];
  const documents = [];

  if (visibility.phyto === true && ukdocsCollectionNeedsPhyto(collection, customer)) {
    documents.push(...(collection?.documents?.phyto_files || []).map((document) => ({ document, kind: "phyto" })));
  }
  if (visibility.export_extra === true && collection?.documents?.export_extra) {
    documents.push({ document: collection.documents.export_extra, kind: "export_extra" });
  }
  if (visibility.inspection_list === true && collection?.documents?.inspection_list) {
    documents.push({ document: collection.documents.inspection_list, kind: "inspection_list" });
  }
  if (visibility.locations_file === true && collection?.documents?.locations_file) {
    documents.push({ document: collection.documents.locations_file, kind: "locations_file" });
  }
  for (const generatedFile of generatedFiles) {
    if (generatedFile?.document_kind === "invoice" && visibility.generated_invoice !== true) {
      continue;
    }
    if (generatedFile?.document_kind === "export" && visibility.generated_export !== true) {
      continue;
    }
    if (!generatedFile?.document_kind && visibility.generated_invoice !== true && visibility.generated_export !== true) {
      continue;
    }
    if (generatedFile?.document_kind === "invoice" && !isPdfUkdocsDocument(generatedFile)) {
      continue;
    }
    documents.push({ document: generatedFile, kind: "generated" });
  }

  for (const item of documents) {
    const document = item.document;
    const resolvedPath = path.resolve(ukdocsPrintDocumentPath(document));
    if (!resolvedPath.startsWith(path.resolve(ukdocsPrintFilesDir)) || !existsSync(resolvedPath)) {
      continue;
    }
    const contentBase64 = await fs.readFile(resolvedPath, "base64");
    attachments.push({
      file_name: path.basename(document.original_name || resolvedPath),
      mime_type: document.mime_type || guessMimeType(document.original_name || resolvedPath),
      content_base64: contentBase64,
    });
  }
  return attachments;
}

function buildUkdocsPrintReadyTemplateContext(collection, requirements) {
  return {
    customer_name: requirements.customer?.customer_name || collection.customer_name || collection.city_name || "-",
    shipment_reference: collection.shipment_reference || "-",
    shipment_date: collection.shipment_date || "-",
    city: collection.city_name || "-",
    hub_code: collection.hub_code || "-",
    reference_connect: collection.reference_connect || "-",
    invoice_numbers: collection.invoice_numbers || "-",
    truck_number: collection.truck_number || "-",
    trailer_number: collection.trailer_number || "-",
    border_crossing: collection.border_crossing || "-",
    pd_form: collection.pd_form || "-",
    re_export: collection.re_export || "-",
    pd_type: collection.pd_type || "-",
    pd_code: collection.pd_code || "-",
    notes: collection.notes || "",
  };
}

function isPdfUkdocsDocument(document) {
  const fileName = String(document?.original_name || document?.storage_name || "").trim().toLowerCase();
  const mimeType = String(document?.mime_type || "").trim().toLowerCase();
  return fileName.endsWith(".pdf") || mimeType === "application/pdf";
}

function applyUkdocsReadyTemplate(template, context) {
  return String(template || "").replace(/\{([a-z_]+)\}/gi, (match, key) => {
    const normalizedKey = String(key || "").toLowerCase();
    return Object.prototype.hasOwnProperty.call(context, normalizedKey) ? String(context[normalizedKey] || "") : match;
  });
}

function buildUkdocsPrintReadyEmail(collection, requirements) {
  const context = buildUkdocsPrintReadyTemplateContext(collection, requirements);
  if (String(requirements.customer?.ready_email_body || "").trim()) {
    return applyUkdocsReadyTemplate(requirements.customer.ready_email_body, context).trim();
  }
  return [
    "UKdocs shipment papers are ready.",
    "",
    `Customer: ${context.customer_name}`,
    `Shipment date: ${context.shipment_date}`,
    `City: ${context.city}`,
    `Hub code: ${context.hub_code}`,
    `Reference connect: ${context.reference_connect}`,
    `Invoices: ${context.invoice_numbers}`,
    `Truck: ${context.truck_number}`,
    `Trailer: ${context.trailer_number}`,
    "",
    `Border crossing: ${context.border_crossing}`,
    `PD form: ${context.pd_form}`,
    `Re-export: ${context.re_export}`,
    `PD type: ${context.pd_type}`,
    `PD code: ${context.pd_code}`,
    "",
    context.notes ? `Notes: ${context.notes}` : "",
  ].filter(Boolean).join("\n");
}

async function sendUkdocsPrintReadyEmail(collection, customers, settings) {
  const recipients = normalizeEmailRecipients(settings.email_recipients);
  if (!recipients.length) {
    return { ok: false, recipients: [], error: "No email recipients configured" };
  }
  const requirements = getUkdocsPrintCollectionRequirements(collection, customers);
  if (!requirements.complete) {
    return { ok: false, recipients, error: `Still missing: ${requirements.missing.join(", ")}` };
  }
  const attachments = await ukdocsPrintCollectionAttachments(collection, requirements.customer, "ukdocsprint");
  const context = buildUkdocsPrintReadyTemplateContext(collection, requirements);
  const subject = String(requirements.customer?.ready_email_subject || "").trim()
    ? applyUkdocsReadyTemplate(requirements.customer.ready_email_subject, context).trim()
    : `UKdocs ready | ${context.customer_name} | ${context.shipment_date}`;
  await runPythonBridge(
    ["email-send"],
    JSON.stringify({
      recipients,
      subject,
      body: buildUkdocsPrintReadyEmail(collection, requirements),
      attachments,
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

async function ukdocsCsiCollectionAttachments(collection, options = {}) {
  const includeSupportDocs = options.include_support_docs !== false;
  const attachments = [];
  const documents = [];
  const generatedFiles = Array.isArray(collection?.documents?.generated_files) ? collection.documents.generated_files : [];

  documents.push(
    ...generatedFiles
      .filter((file) => file?.document_kind === "invoice" && isPdfUkdocsDocument(file))
      .map((document) => ({ document, kind: "generated_invoice_pdf" })),
  );

  const generatedExport = generatedFiles.find((file) => file?.document_kind === "export");
  if (generatedExport) {
    documents.push({ document: generatedExport, kind: "generated_export" });
  }

  if (includeSupportDocs) {
    documents.push(
      ...(collection?.documents?.temp_phyto_files || []).map((document) => ({ document, kind: "temp_phyto" })),
    );
    if (collection?.documents?.temp_phyto_plants_file) {
      documents.push({ document: collection.documents.temp_phyto_plants_file, kind: "temp_phyto_plants_file" });
    }

    if (collection?.documents?.ipaffs_file) {
      documents.push({ document: collection.documents.ipaffs_file, kind: "ipaffs_file" });
    }
    if (collection?.documents?.ipaffs_plants_file) {
      documents.push({ document: collection.documents.ipaffs_plants_file, kind: "ipaffs_plants_file" });
    }
  }

  for (const item of documents) {
    const document = item.document;
    const resolvedPath = path.resolve(ukdocsPrintDocumentPath(document));
    if (!resolvedPath.startsWith(path.resolve(ukdocsPrintFilesDir)) || !existsSync(resolvedPath)) {
      continue;
    }
    const contentBase64 = await fs.readFile(resolvedPath, "base64");
    attachments.push({
      file_name: path.basename(document.original_name || resolvedPath),
      mime_type: document.mime_type || guessMimeType(document.original_name || resolvedPath),
      content_base64: contentBase64,
    });
  }

  return attachments;
}

function buildUkdocsCsiTemplateContext(collection, customer = null) {
  return buildUkdocsPrintReadyTemplateContext(collection, {
    customer,
  });
}

function buildUkdocsCsiSuccessEmail(collection, customer = null, options = {}) {
  const context = buildUkdocsCsiTemplateContext(collection, customer);
  const includeSupportDocs = options.include_support_docs !== false;
  if (String(customer?.csi_email_body || "").trim()) {
    return applyUkdocsReadyTemplate(customer.csi_email_body, context).trim();
  }
  return [
    "CSI audit completed successfully.",
    "",
    `Customer: ${context.customer_name}`,
    `Shipment reference: ${context.shipment_reference}`,
    `Shipment date: ${context.shipment_date}`,
    `City: ${context.city}`,
    `Reference connect: ${context.reference_connect}`,
    `Invoices: ${context.invoice_numbers}`,
    `Truck: ${context.truck_number}`,
    `Trailer: ${context.trailer_number}`,
    "",
    "Attached files:",
    "- Generated invoice PDF files",
    "- Generated export file",
    includeSupportDocs ? "- Temporary phyto PDF files" : "",
    includeSupportDocs ? "- Temporary phyto plants PDF file" : "",
    includeSupportDocs ? "- IPAFFS file(s)" : "",
  ].join("\n");
}

async function sendUkdocsCsiSuccessEmail(collection, customers, settings, options = {}) {
  const customer = ukdocsPrintCollectionCustomer(collection, customers);
  const recipients = normalizeEmailRecipients(customer?.csi_email_recipients);
  const noPdNeeded = isUkdocsNoPdNeeded(collection);
  const includeSupportDocs = options.include_support_docs !== false && !noPdNeeded;
  if (!recipients.length) {
    return { ok: false, recipients: [], error: "No CSI email recipients configured for this customer" };
  }
  if (!settings?.smtp_host || !settings?.smtp_username || !settings?.smtp_password || !settings?.smtp_from) {
    return { ok: false, recipients, error: "SMTP is not fully configured" };
  }
  const attachments = await ukdocsCsiCollectionAttachments(collection, { include_support_docs: includeSupportDocs });
  if (!attachments.length) {
    return { ok: false, recipients, error: "No CSI attachments found to send" };
  }
  const generatedInvoicePdfCount = (collection?.documents?.generated_files || []).filter((file) => file?.document_kind === "invoice" && isPdfUkdocsDocument(file)).length;
  if (!generatedInvoicePdfCount) {
    return { ok: false, recipients, error: "No generated invoice PDF files found for CSI email" };
  }
  if (!(collection?.documents?.generated_files || []).some((file) => file?.document_kind === "export")) {
    return { ok: false, recipients, error: "No generated export file found for CSI email" };
  }
  if (includeSupportDocs && !(collection?.documents?.temp_phyto_files || []).length && !collection?.documents?.temp_phyto_plants_file?.storage_name) {
    return { ok: false, recipients, error: "No temporary phyto PDF files found for CSI email" };
  }
  if (includeSupportDocs && !collection?.documents?.ipaffs_file?.storage_name && !collection?.documents?.ipaffs_plants_file?.storage_name) {
    return { ok: false, recipients, error: "No IPAFFS file found for CSI email" };
  }
  const context = buildUkdocsCsiTemplateContext(collection, customer);
  const subject = String(customer?.csi_email_subject || "").trim()
    ? applyUkdocsReadyTemplate(customer.csi_email_subject, context).trim()
    : `CSI OK | ${context.customer_name} | ${context.shipment_date}`;
  await runPythonBridge(
    ["email-send"],
    JSON.stringify({
      recipients,
      subject,
      body: buildUkdocsCsiSuccessEmail(collection, customer, { include_support_docs: includeSupportDocs }),
      attachments,
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

function ukdocsGmailRedirectUri(req) {
  return `${publicBaseUrl(req)}/api/ukdocs-print/gmail/callback`;
}

function googleAuthUrl(settings, req, redirectUri, scopes) {
  const params = new URLSearchParams({
    client_id: settings.cmr_google_client_id,
    redirect_uri: redirectUri,
    response_type: "code",
    scope: scopes.join(" "),
    access_type: "offline",
    prompt: "consent",
  });
  return `https://accounts.google.com/o/oauth2/v2/auth?${params.toString()}`;
}

function cmrGoogleAuthUrl(settings, req) {
  return googleAuthUrl(
    settings,
    req,
    cmrGoogleRedirectUri(req),
    [
      "https://www.googleapis.com/auth/drive",
      "https://www.googleapis.com/auth/userinfo.email",
      "openid",
    ],
  );
}

function ukdocsGmailAuthUrl(settings, req) {
  return googleAuthUrl(
    settings,
    req,
    ukdocsGmailRedirectUri(req),
    [
      "https://www.googleapis.com/auth/gmail.readonly",
      "https://www.googleapis.com/auth/userinfo.email",
      "openid",
    ],
  );
}

async function exchangeGoogleAuthCode(settings, code, redirectUri) {
  const response = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      code,
      client_id: settings.cmr_google_client_id,
      client_secret: settings.cmr_google_client_secret,
      redirect_uri: redirectUri,
      grant_type: "authorization_code",
    }),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(payload.error_description || payload.error || `Google token exchange failed with ${response.status}`);
  }
  return payload;
}

async function refreshGoogleAccessToken(settings, refreshToken) {
  if (!settings.cmr_google_client_id || !settings.cmr_google_client_secret || !refreshToken) {
    throw new Error("Google OAuth is not fully configured");
  }
  const response = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: settings.cmr_google_client_id,
      client_secret: settings.cmr_google_client_secret,
      refresh_token: refreshToken,
      grant_type: "refresh_token",
    }),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok || !payload.access_token) {
    throw new Error(payload.error_description || payload.error || `Google token refresh failed with ${response.status}`);
  }
  return String(payload.access_token);
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

  if (url.pathname === "/api/public/clock/employees" && req.method === "GET") {
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

  if (url.pathname === "/api/public/clock/scan" && req.method === "POST") {
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
    const record = createClockRecord(employee, direction, actionDate, actionTime, "scanner", "public-kiosk");
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
    sendJson(res, 201, {
      record,
      records: filterClockRecords(savedRecords, actionDate),
      sessions: clockSessionRows(savedRecords.filter((item) => item.action_date === actionDate)).map((session) => ({ in_record: session.inRecord, out_record: session.outRecord, row: session.row })),
    });
    return;
  }

  if (url.pathname === "/api/llm/agent/heartbeat" && req.method === "POST") {
    if (!isDatabaseEnabled()) {
      sendJson(res, 503, { error: "Database is not enabled" });
      return;
    }
    if (!llmPollerEnabled()) {
      sendJson(res, 503, { error: "SHADOW_LLM_POLLER_API_KEY is not configured" });
      return;
    }
    const apiKey = readAgentApiKey(req);
    if (!apiKey || apiKey !== llmPollerApiKey) {
      sendPollerUnauthorized(res);
      return;
    }
    const body = await readRequestJson(req);
    const agentName = String(body.agent_name || "").trim();
    if (!agentName) {
      sendJson(res, 400, { error: "agent_name is required" });
      return;
    }
    const agent = await upsertLlmAgentHeartbeat({
      agent_name: agentName,
      pc_name: body.pc_name,
      model_name: body.model_name,
      version: body.version,
      status: body.status || "online",
      capabilities: Array.isArray(body.capabilities) ? body.capabilities : [],
      meta: body.meta && typeof body.meta === "object" ? body.meta : {},
    }, apiKey);
    sendJson(res, 200, { ok: true, agent, server_time: new Date().toISOString() });
    return;
  }

  if (url.pathname === "/api/llm/agent/poll" && req.method === "POST") {
    if (!isDatabaseEnabled()) {
      sendJson(res, 503, { error: "Database is not enabled" });
      return;
    }
    if (!llmPollerEnabled()) {
      sendJson(res, 503, { error: "SHADOW_LLM_POLLER_API_KEY is not configured" });
      return;
    }
    const apiKey = readAgentApiKey(req);
    if (!apiKey || apiKey !== llmPollerApiKey) {
      sendPollerUnauthorized(res);
      return;
    }
    const body = await readRequestJson(req);
    const agentName = String(body.agent_name || "").trim();
    if (!agentName) {
      sendJson(res, 400, { error: "agent_name is required" });
      return;
    }
    await upsertLlmAgentHeartbeat({
      agent_name: agentName,
      pc_name: body.pc_name,
      model_name: body.model_name,
      version: body.version,
      status: body.status || "online",
      capabilities: Array.isArray(body.capabilities) ? body.capabilities : [],
      meta: body.meta && typeof body.meta === "object" ? body.meta : {},
    }, apiKey);
    const result = await claimNextLlmJob(agentName, apiKey, {
      agent_status: body.agent_status || "idle",
    });
    sendJson(res, 200, {
      ok: true,
      server_time: new Date().toISOString(),
      agent: result?.agent || null,
      job: result?.job || null,
    });
    return;
  }

  if (url.pathname.startsWith("/api/llm/jobs/") && url.pathname.endsWith("/result") && req.method === "POST") {
    if (!isDatabaseEnabled()) {
      sendJson(res, 503, { error: "Database is not enabled" });
      return;
    }
    if (!llmPollerEnabled()) {
      sendJson(res, 503, { error: "SHADOW_LLM_POLLER_API_KEY is not configured" });
      return;
    }
    const apiKey = readAgentApiKey(req);
    if (!apiKey || apiKey !== llmPollerApiKey) {
      sendPollerUnauthorized(res);
      return;
    }
    const jobId = decodeURIComponent(url.pathname.slice("/api/llm/jobs/".length, -"/result".length));
    const body = await readRequestJson(req);
    const agentName = String(body.agent_name || "").trim();
    if (!jobId || !agentName) {
      sendJson(res, 400, { error: "job id and agent_name are required" });
      return;
    }
    const job = await completeLlmJob(jobId, agentName, body.result_json && typeof body.result_json === "object" ? body.result_json : {});
    if (!job) {
      sendJson(res, 404, { error: "LLM job not found for this agent" });
      return;
    }
    if (job.job_type === "excel_to_pdf" && job.collection_id) {
      await saveUkdocsGeneratedInvoicePdfResult(job);
    }
    if (job.job_type === "ukdocs_csi_audit" && job.collection_id) {
      const csiGroupId = String(job?.payload_json?.csi_group_id || "").trim();
      if (csiGroupId) {
        const groupJobs = await getUkdocsCsiGroupJobs(job.collection_id, csiGroupId);
        const allDone = groupJobs.length > 0 && groupJobs.every((item) => item.status === "done");
        if (!allDone) {
          await updateUkdocsCsiReport(job.collection_id, {
            status: "running",
            started_at: groupJobs.find((item) => item.claimed_at)?.claimed_at || job.claimed_at || "",
            error: "",
            summary: "CSI audit is still processing the temporary phyto documents.",
          });
        } else {
          const finalReport = buildUkdocsCsiReportFromJobResults(groupJobs.map(parseUkdocsCsiAuditJobResult));
          await updateUkdocsCsiReport(job.collection_id, {
            ...finalReport,
            started_at: groupJobs.find((item) => item.claimed_at)?.claimed_at || job.claimed_at || "",
            completed_at: groupJobs.reduce((latest, item) => {
              const candidate = String(item?.finished_at || "").trim();
              return candidate > latest ? candidate : latest;
            }, ""),
          });
        }
      } else {
        const finalReport = buildUkdocsCsiReportFromJobResults([parseUkdocsCsiAuditJobResult(job)]);
        await updateUkdocsCsiReport(job.collection_id, {
          ...finalReport,
          started_at: job.claimed_at || "",
          completed_at: job.finished_at || new Date().toISOString(),
        });
      }
      await updateUkdocsCsiEmailResult(job.collection_id, {
        ok: false,
        recipients: [],
        sent_at: "",
        error: "",
      });
    }
    await upsertLlmAgentHeartbeat({
      agent_name: agentName,
      pc_name: body.pc_name,
      model_name: body.model_name,
      version: body.version,
      status: body.status || "idle",
      capabilities: Array.isArray(body.capabilities) ? body.capabilities : [],
      meta: body.meta && typeof body.meta === "object" ? body.meta : {},
    }, apiKey);
    sendJson(res, 200, { ok: true, job });
    return;
  }

  if (url.pathname.startsWith("/api/llm/jobs/") && url.pathname.endsWith("/fail") && req.method === "POST") {
    if (!isDatabaseEnabled()) {
      sendJson(res, 503, { error: "Database is not enabled" });
      return;
    }
    if (!llmPollerEnabled()) {
      sendJson(res, 503, { error: "SHADOW_LLM_POLLER_API_KEY is not configured" });
      return;
    }
    const apiKey = readAgentApiKey(req);
    if (!apiKey || apiKey !== llmPollerApiKey) {
      sendPollerUnauthorized(res);
      return;
    }
    const jobId = decodeURIComponent(url.pathname.slice("/api/llm/jobs/".length, -"/fail".length));
    const body = await readRequestJson(req);
    const agentName = String(body.agent_name || "").trim();
    if (!jobId || !agentName) {
      sendJson(res, 400, { error: "job id and agent_name are required" });
      return;
    }
    const job = await failLlmJob(jobId, agentName, body.error_text || "Job failed", body.allow_retry === true);
    if (!job) {
      sendJson(res, 404, { error: "LLM job not found for this agent" });
      return;
    }
    if (job.job_type === "ukdocs_csi_audit" && job.collection_id) {
      const label = String(job?.payload_json?.csi_document_label || "").trim();
      await updateUkdocsCsiReport(job.collection_id, {
        status: "failed",
        started_at: job.claimed_at || "",
        completed_at: job.finished_at || new Date().toISOString(),
        error: String(body.error_text || job.error_text || "CSI audit failed").trim(),
        summary: "CSI audit failed.",
        notes: label ? [`Failed while processing ${label}.`] : [],
      });
    }
    await upsertLlmAgentHeartbeat({
      agent_name: agentName,
      pc_name: body.pc_name,
      model_name: body.model_name,
      version: body.version,
      status: body.status || "idle",
      capabilities: Array.isArray(body.capabilities) ? body.capabilities : [],
      meta: body.meta && typeof body.meta === "object" ? body.meta : {},
    }, apiKey);
    sendJson(res, 200, { ok: true, job });
    return;
  }

  const requestUser = await getRequestUser(req);
  if (!requestUser) {
    sendUnauthorized(res);
    return;
  }

  if (url.pathname === "/api/expedition-stickers" && req.method === "GET") {
    if (!requirePermission(res, requestUser, PERMISSIONS.EXPEDITION_STICKERS_VIEW)) {
      return;
    }

    try {
      const settings = await readFustSettings();
      const state = await readExpeditionStickerState();
      const payload = {
        planning_file: state.planning_file,
        split_file: state.split_file,
        sheet_source: {
          spreadsheet_id: String(settings.hal_locations_spreadsheet_id || settings.spreadsheet_id || "").trim(),
          sheet_name: String(settings.hal_locations_sheet_name || "ERP_PASTE").trim() || "ERP_PASTE",
        },
      };

      if (state.planning_file) {
        const planningPath = expeditionStickerFilePath(state.planning_file);
        if (planningPath && existsSync(planningPath)) {
          payload.planning_summary = await inspectExpeditionStickerSource("planning", planningPath);
        }
      }
      if (state.split_file) {
        const splitPath = expeditionStickerFilePath(state.split_file);
        if (splitPath && existsSync(splitPath)) {
          payload.split_summary = await inspectExpeditionStickerSource("split", splitPath);
        }
      }

      sendJson(res, 200, payload);
    } catch (error) {
      sendJson(res, 400, { error: error instanceof Error ? error.message : String(error) });
    }
    return;
  }

  if (url.pathname === "/api/llm/status" && req.method === "GET") {
    if (!requirePermission(res, requestUser, PERMISSIONS.SETTINGS_MANAGE)) {
      return;
    }
    const snapshot = await getLlmQueueSnapshot();
    sendJson(res, 200, {
      ok: true,
      poller_enabled: llmPollerEnabled(),
      agent_key_configured: llmPollerEnabled(),
      snapshot,
    });
    return;
  }

  if (url.pathname === "/api/llm/jobs" && req.method === "POST") {
    if (!requirePermission(res, requestUser, PERMISSIONS.SETTINGS_MANAGE)) {
      return;
    }
    if (!isDatabaseEnabled()) {
      sendJson(res, 503, { error: "Database is not enabled" });
      return;
    }
    const body = await readRequestJson(req);
    const jobType = String(body.job_type || "").trim();
    if (!jobType) {
      sendJson(res, 400, { error: "job_type is required" });
      return;
    }
    const job = await createLlmJob({
      job_type: jobType,
      created_by: requestUser.username,
      shipment_id: body.shipment_id,
      collection_id: body.collection_id,
      document_kind: body.document_kind,
      priority: body.priority,
      max_attempts: body.max_attempts,
      payload_json: body.payload_json && typeof body.payload_json === "object" ? body.payload_json : {},
    });
    sendJson(res, 201, { ok: true, job });
    return;
  }

  if (url.pathname === "/api/expedition-stickers/upload" && req.method === "POST") {
    if (!requirePermission(res, requestUser, PERMISSIONS.EXPEDITION_STICKERS_VIEW)) {
      return;
    }

    try {
      const body = await readRequestJson(req, 50 * 1024 * 1024);
      const currentState = await readExpeditionStickerState();
      const nextState = { ...currentState };

      if (body?.planning_file?.content_base64) {
        nextState.planning_file = await saveExpeditionStickerUpload("planning", body.planning_file, requestUser);
      }
      if (body?.split_file?.content_base64) {
        nextState.split_file = await saveExpeditionStickerUpload("split", body.split_file, requestUser);
      }
      if (!body?.planning_file?.content_base64 && !body?.split_file?.content_base64) {
        sendJson(res, 400, { error: "Upload a planning file or a split file first" });
        return;
      }

      await writeExpeditionStickerState(nextState);
      const responsePayload = {
        planning_file: nextState.planning_file,
        split_file: nextState.split_file,
      };

      if (nextState.planning_file) {
        responsePayload.planning_summary = await inspectExpeditionStickerSource("planning", expeditionStickerFilePath(nextState.planning_file));
      }
      if (nextState.split_file) {
        responsePayload.split_summary = await inspectExpeditionStickerSource("split", expeditionStickerFilePath(nextState.split_file));
      }

      sendJson(res, 200, responsePayload);
    } catch (error) {
      sendJson(res, 400, { error: error instanceof Error ? error.message : String(error) });
    }
    return;
  }

  if (url.pathname.match(/^\/api\/expedition-stickers\/source\/(planning|split)$/) && req.method === "DELETE") {
    if (!requirePermission(res, requestUser, PERMISSIONS.EXPEDITION_STICKERS_VIEW)) {
      return;
    }

    try {
      const kind = url.pathname.split("/").pop();
      const nextState = await deleteExpeditionStickerUpload(kind);
      const responsePayload = {
        planning_file: nextState.planning_file,
        split_file: nextState.split_file,
        planning_summary: null,
        split_summary: null,
      };

      if (nextState.planning_file) {
        responsePayload.planning_summary = await inspectExpeditionStickerSource("planning", expeditionStickerFilePath(nextState.planning_file));
      }
      if (nextState.split_file) {
        responsePayload.split_summary = await inspectExpeditionStickerSource("split", expeditionStickerFilePath(nextState.split_file));
      }

      sendJson(res, 200, responsePayload);
    } catch (error) {
      sendJson(res, 400, { error: error instanceof Error ? error.message : String(error) });
    }
    return;
  }

  if (url.pathname === "/api/expedition-stickers/load-sheet" && req.method === "POST") {
    if (!requirePermission(res, requestUser, PERMISSIONS.EXPEDITION_STICKERS_VIEW)) {
      return;
    }

    let session = null;
    try {
      const settings = await readFustSettings();
      const spreadsheetId = String(settings.hal_locations_spreadsheet_id || settings.spreadsheet_id || "").trim();
      const sheetName = String(settings.hal_locations_sheet_name || "ERP_PASTE").trim() || "ERP_PASTE";
      if (!spreadsheetId) {
        sendJson(res, 400, { error: "Set a Hal Locations spreadsheet ID in Settings first" });
        return;
      }

      const rows = await loadSheetRows(spreadsheetId, sheetName);
      session = await createHalLocationSheetSession(rows, { spreadsheet_id: spreadsheetId, sheet_name: sheetName });
      const output = await runHalLocationsWorker([
        "inspect",
        "--input",
        session.file_path,
      ]);
      const payload = JSON.parse(output.toString("utf8"));
      sendJson(res, 200, {
        id: session.id,
        locPrefixes: Array.isArray(payload.locPrefixes) ? payload.locPrefixes : [],
        custPrefixes: Array.isArray(payload.custPrefixes) ? payload.custPrefixes : [],
        custByLoc: payload.custByLoc && typeof payload.custByLoc === "object" ? payload.custByLoc : {},
        totalRows: Number(payload.totalRows || 0),
        source: {
          type: "sheet",
          spreadsheet_id: spreadsheetId,
          sheet_name: sheetName,
        },
      });
    } catch (error) {
      if (session?.id) {
        halLocationSessions.delete(session.id);
      }
      if (session?.dir_path) {
        await fs.rm(session.dir_path, { recursive: true, force: true }).catch(() => {});
      }
      sendJson(res, 400, { error: error instanceof Error ? error.message : String(error) });
    }
    return;
  }

  if (url.pathname === "/api/expedition-stickers/generate" && req.method === "POST") {
    if (!requirePermission(res, requestUser, PERMISSIONS.EXPEDITION_STICKERS_VIEW)) {
      return;
    }

    const body = await readRequestJson(req, 2 * 1024 * 1024);
    const session = await getHalLocationSession(body.id);
    if (!session) {
      sendJson(res, 404, { error: "Session expired. Load ERP_PASTE again first." });
      return;
    }

    const state = await readExpeditionStickerState();
    const planningPath = state.planning_file ? expeditionStickerFilePath(state.planning_file) : "";
    const splitPath = state.split_file ? expeditionStickerFilePath(state.split_file) : "";
    if (!planningPath && !splitPath) {
      sendJson(res, 400, { error: "Upload and save a planning file or split file first" });
      return;
    }

    const outputDir = path.join(session.dir_path, `expedition-${Date.now()}`);
    try {
      const args = [
        "generate",
        "--hal-input",
        session.file_path,
        "--output-dir",
        outputDir,
      ];
      if (planningPath && existsSync(planningPath)) {
        args.push("--planning-input", planningPath);
      }
      if (splitPath && existsSync(splitPath)) {
        args.push("--split-input", splitPath);
      }

      const output = await runExpeditionStickerWorker(args);
      const payload = JSON.parse(output.toString("utf8"));
      const files = await Promise.all(
        (Array.isArray(payload.files) ? payload.files : []).map(async (file) => {
          const fileBuffer = await fs.readFile(String(file.path || ""));
          return {
            name: String(file.name || "stickers.pdf"),
            mime_type: "application/pdf",
            content_base64: fileBuffer.toString("base64"),
            split: file.split ?? null,
            customer_count: Number(file.customer_count || 0),
            sticker_count: Number(file.sticker_count || 0),
          };
        }),
      );

      sendJson(res, 200, {
        summary: {
          hal_customer_count: Number(payload.hal_customer_count || 0),
          combined_row_count: Number(payload.combined_row_count || 0),
          missing_locations: Array.isArray(payload.missing_locations) ? payload.missing_locations : [],
        },
        files,
      });
    } catch (error) {
      sendJson(res, 400, { error: error instanceof Error ? error.message : String(error) });
    } finally {
      await fs.rm(outputDir, { recursive: true, force: true }).catch(() => {});
    }
    return;
  }

  if (url.pathname === "/api/hal-locations/inspect" && req.method === "POST") {
    if (!requirePermission(res, requestUser, PERMISSIONS.HAL_LOCATIONS_VIEW)) {
      return;
    }

    const body = await readRequestJson(req, 25 * 1024 * 1024);
    let session = null;
    try {
      session = await createHalLocationSession(body.file || {});
      const output = await runHalLocationsWorker([
        "inspect",
        "--input",
        session.file_path,
      ]);
      const payload = JSON.parse(output.toString("utf8"));
      sendJson(res, 200, {
        id: session.id,
        locPrefixes: Array.isArray(payload.locPrefixes) ? payload.locPrefixes : [],
        custPrefixes: Array.isArray(payload.custPrefixes) ? payload.custPrefixes : [],
        custByLoc: payload.custByLoc && typeof payload.custByLoc === "object" ? payload.custByLoc : {},
        totalRows: Number(payload.totalRows || 0),
        source: { type: "upload", file_name: session.file_name },
      });
    } catch (error) {
      if (session?.id) {
        halLocationSessions.delete(session.id);
      }
      if (session?.dir_path) {
        await fs.rm(session.dir_path, { recursive: true, force: true }).catch(() => {});
      }
      sendJson(res, 400, { error: error instanceof Error ? error.message : String(error) });
    }
    return;
  }

  if (url.pathname === "/api/dag-foutjes/app" && req.method === "GET") {
    if (!requirePermission(res, requestUser, PERMISSIONS.EXPEDITION_STICKERS_VIEW)) {
      return;
    }
    const dagFoutjesHtmlPath = resolveDagFoutjesHtmlPath();
    if (!dagFoutjesHtmlPath) {
      sendText(res, 404, `Dag Foutjes app not found. Checked: ${dagFoutjesHtmlPathCandidates.join(" | ")}`);
      return;
    }
    const html = await fs.readFile(dagFoutjesHtmlPath, "utf8");
    const withBridge = html.includes("window.storage")
      ? html.replace("<script>", `${dagFoutjesBridgeScript()}\n<script>`)
      : `${dagFoutjesBridgeScript()}\n${html}`;
    res.writeHead(200, {
      "content-type": "text/html; charset=utf-8",
      "cache-control": "no-store",
    });
    res.end(withBridge);
    return;
  }

  if (url.pathname === "/api/bunches/app" && req.method === "GET") {
    if (!requirePermission(res, requestUser, PERMISSIONS.BUNCHES_VIEW)) {
      return;
    }
    const html = await fs.readFile(bunchesAppHtmlPath, "utf8");
    res.writeHead(200, {
      "content-type": "text/html; charset=utf-8",
      "cache-control": "no-store",
    });
    res.end(html);
    return;
  }

  if (url.pathname === "/api/bunches/state" && req.method === "GET") {
    if (!requirePermission(res, requestUser, PERMISSIONS.BUNCHES_VIEW)) {
      return;
    }
    sendJson(res, 200, await bunchesService.getAppState());
    return;
  }

  if (url.pathname === "/api/bunches/process" && req.method === "POST") {
    if (!requirePermission(res, requestUser, PERMISSIONS.BUNCHES_VIEW)) {
      return;
    }
    const body = await readRequestJson(req, 10 * 1024 * 1024);
    try {
      const run = await bunchesService.processImport({
        pasteText: body.paste_text || "",
        vertrekDatum: body.vertrek_datum || "",
        label: body.label || "",
        user: requestUser?.username || "unknown",
      });
      sendJson(res, 200, { run });
    } catch (error) {
      sendJson(res, 400, { error: error instanceof Error ? error.message : String(error) });
    }
    return;
  }

  if (url.pathname.match(/^\/api\/bunches\/runs\/\d+\/date$/) && req.method === "PATCH") {
    if (!requirePermission(res, requestUser, PERMISSIONS.BUNCHES_VIEW)) {
      return;
    }
    const runId = Number(url.pathname.split("/")[4]);
    const body = await readRequestJson(req);
    try {
      sendJson(res, 200, { run: await bunchesService.updateRunDate(runId, body.vertrek_datum || "") });
    } catch (error) {
      sendJson(res, 400, { error: error instanceof Error ? error.message : String(error) });
    }
    return;
  }

  if (url.pathname.match(/^\/api\/bunches\/runs\/\d+\/label$/) && req.method === "PATCH") {
    if (!requirePermission(res, requestUser, PERMISSIONS.BUNCHES_VIEW)) {
      return;
    }
    const runId = Number(url.pathname.split("/")[4]);
    const body = await readRequestJson(req);
    try {
      sendJson(res, 200, { run: await bunchesService.updateRunLabel(runId, body.label || "") });
    } catch (error) {
      sendJson(res, 400, { error: error instanceof Error ? error.message : String(error) });
    }
    return;
  }

  if (url.pathname.match(/^\/api\/bunches\/runs\/\d+$/) && req.method === "DELETE") {
    if (!requirePermission(res, requestUser, PERMISSIONS.BUNCHES_VIEW)) {
      return;
    }
    const runId = Number(url.pathname.split("/")[4]);
    try {
      await bunchesService.deleteRun(runId);
      sendJson(res, 200, { ok: true });
    } catch (error) {
      sendJson(res, 400, { error: error instanceof Error ? error.message : String(error) });
    }
    return;
  }

  if (url.pathname === "/api/bunches/articles" && req.method === "POST") {
    if (!requirePermission(res, requestUser, PERMISSIONS.BUNCHES_VIEW)) {
      return;
    }
    const body = await readRequestJson(req);
    try {
      sendJson(res, 200, { article: await bunchesService.upsertArticle(body) });
    } catch (error) {
      sendJson(res, 400, { error: error instanceof Error ? error.message : String(error) });
    }
    return;
  }

  if (url.pathname === "/api/bunches/articles/bulk-zonder-tak" && req.method === "POST") {
    if (!requirePermission(res, requestUser, PERMISSIONS.BUNCHES_VIEW)) {
      return;
    }
    const body = await readRequestJson(req);
    const broncodes = String(body.broncodes || "")
      .split(/[\s,;]+/)
      .map((value) => Number(value))
      .filter((value) => Number.isFinite(value));
    try {
      sendJson(res, 200, await bunchesService.bulkSetZonderTak(broncodes, Boolean(body.value)));
    } catch (error) {
      sendJson(res, 400, { error: error instanceof Error ? error.message : String(error) });
    }
    return;
  }

  if (url.pathname.match(/^\/api\/bunches\/articles\/\d+$/) && req.method === "DELETE") {
    if (!requirePermission(res, requestUser, PERMISSIONS.BUNCHES_VIEW)) {
      return;
    }
    const broncode = Number(url.pathname.split("/")[4]);
    try {
      await bunchesService.deactivateArticle(broncode);
      sendJson(res, 200, { ok: true });
    } catch (error) {
      sendJson(res, 400, { error: error instanceof Error ? error.message : String(error) });
    }
    return;
  }

  if (url.pathname === "/api/bunches/ape" && req.method === "POST") {
    if (!requirePermission(res, requestUser, PERMISSIONS.BUNCHES_VIEW)) {
      return;
    }
    const body = await readRequestJson(req);
    try {
      sendJson(res, 200, { entry: await bunchesService.upsertApe(body) });
    } catch (error) {
      sendJson(res, 400, { error: error instanceof Error ? error.message : String(error) });
    }
    return;
  }

  if (url.pathname === "/api/bunches/ape" && req.method === "DELETE") {
    if (!requirePermission(res, requestUser, PERMISSIONS.BUNCHES_VIEW)) {
      return;
    }
    const body = await readRequestJson(req);
    try {
      await bunchesService.deleteApe(String(body.omschrijving || ""));
      sendJson(res, 200, { ok: true });
    } catch (error) {
      sendJson(res, 400, { error: error instanceof Error ? error.message : String(error) });
    }
    return;
  }

  if (url.pathname.match(/^\/api\/bunches\/download\/\d+\/inlezen\.csv$/) && req.method === "GET") {
    if (!requirePermission(res, requestUser, PERMISSIONS.BUNCHES_VIEW)) {
      return;
    }
    const runId = Number(url.pathname.split("/")[4]);
    try {
      const file = await bunchesService.downloadFile(runId, "inlezen");
      res.writeHead(200, {
        "content-type": file.contentType,
        "content-disposition": `attachment; filename="${file.filename}"`,
      });
      res.end(file.body);
    } catch (error) {
      sendText(res, 404, error instanceof Error ? error.message : String(error));
    }
    return;
  }

  if (url.pathname.match(/^\/api\/bunches\/download\/\d+\/yybu\/[^/]+\.csv$/) && req.method === "GET") {
    if (!requirePermission(res, requestUser, PERMISSIONS.BUNCHES_VIEW)) {
      return;
    }
    const parts = url.pathname.split("/");
    const runId = Number(parts[4]);
    const sheet = decodeURIComponent(parts[6].replace(/\.csv$/i, ""));
    try {
      const file = await bunchesService.downloadFile(runId, "yybu", sheet);
      res.writeHead(200, {
        "content-type": file.contentType,
        "content-disposition": `attachment; filename="${file.filename}"`,
      });
      res.end(file.body);
    } catch (error) {
      sendText(res, 404, error instanceof Error ? error.message : String(error));
    }
    return;
  }

  if (url.pathname.match(/^\/api\/bunches\/printlijst\/\d+\/(plast|kraft)(?:\/[^/]+)?$/i) && req.method === "GET") {
    if (!requirePermission(res, requestUser, PERMISSIONS.BUNCHES_VIEW)) {
      return;
    }
    const parts = url.pathname.split("/");
    const runId = Number(parts[4]);
    const hoes = parts[5];
    const tak = parts[6] ? decodeURIComponent(parts[6]) : "";
    try {
      const autoPrint = url.searchParams.get("autoprint") === "1";
      const html = await bunchesService.renderPrintlijst(runId, hoes, tak, { autoPrint });
      res.writeHead(200, { "content-type": "text/html; charset=utf-8" });
      res.end(html);
    } catch (error) {
      sendText(res, 404, error instanceof Error ? error.message : String(error));
    }
    return;
  }

  if (url.pathname.match(/^\/api\/bunches\/printlijst-pdf\/\d+\/(plast|kraft)\/[^/]+\.pdf$/i) && req.method === "GET") {
    if (!requirePermission(res, requestUser, PERMISSIONS.BUNCHES_VIEW)) {
      return;
    }
    const parts = url.pathname.split("/");
    const runId = Number(parts[4]);
    const hoes = parts[5];
    const takToken = parts[6] ? decodeURIComponent(parts[6].replace(/\.pdf$/i, "")) : "";
    const tak = takToken.toLowerCase() === "all" ? "" : takToken;
    try {
      const file = await bunchesService.renderPrintlijstPdf(runId, hoes, tak);
      res.writeHead(200, {
        "content-type": file.contentType,
        "content-disposition": `attachment; filename="${file.filename}"`,
        "cache-control": "no-store",
      });
      res.end(file.body);
    } catch (error) {
      sendText(res, 404, error instanceof Error ? error.message : String(error));
    }
    return;
  }

  if (url.pathname === "/api/dag-foutjes/overview" && req.method === "GET") {
    if (!requirePermission(res, requestUser, PERMISSIONS.FOUTEN_OVERVIEW_VIEW)) {
      return;
    }
    try {
      const payload = await buildDagFoutjesOverviewPayload();
      sendJson(res, 200, payload);
    } catch (error) {
      sendJson(res, 400, { error: error instanceof Error ? error.message : String(error) });
    }
    return;
  }

  if (url.pathname === "/api/dag-foutjes/storage") {
    if (!requirePermission(res, requestUser, PERMISSIONS.EXPEDITION_STICKERS_VIEW)) {
      return;
    }

    if (req.method === "GET") {
      const op = String(url.searchParams.get("op") || "get").trim().toLowerCase();
      let state = await readDagFoutjesState();
      if (op === "get" && String(url.searchParams.get("key") || "") === "people") {
        try {
          const merged = await mergeDagFoutjesPeopleFromClock(state);
          state = merged.state;
          if (merged.changed) {
            await writeDagFoutjesState(state);
          }
        } catch {
        }
      }
      if (op === "list") {
        const prefix = String(url.searchParams.get("prefix") || "");
        const keys = Object.keys(state.shared).filter((key) => key.startsWith(prefix)).sort();
        sendJson(res, 200, { keys });
        return;
      }
      const key = String(url.searchParams.get("key") || "");
      if (!key) {
        sendJson(res, 400, { error: "key is required" });
        return;
      }
      const found = Object.prototype.hasOwnProperty.call(state.shared, key);
      sendJson(res, 200, {
        found,
        key,
        value: found ? String(state.shared[key] ?? "") : "",
      });
      return;
    }

    if (req.method === "POST") {
      const body = await readRequestJson(req, 2 * 1024 * 1024);
      const op = String(body?.op || "set").trim().toLowerCase();
      if (op !== "set") {
        sendJson(res, 400, { error: "Unsupported operation" });
        return;
      }
      const key = String(body?.key || "");
      if (!key) {
        sendJson(res, 400, { error: "key is required" });
        return;
      }
      const state = await readDagFoutjesState();
      state.shared[key] = String(body?.value ?? "");
      await writeDagFoutjesState(state);
      sendJson(res, 200, { ok: true, key });
      return;
    }

    sendJson(res, 405, { error: "Method not allowed" });
    return;
  }

  if (url.pathname === "/api/dag-foutjes/sheet-sync" && req.method === "POST") {
    if (!requirePermission(res, requestUser, PERMISSIONS.EXPEDITION_STICKERS_VIEW)) {
      return;
    }

    const body = await readRequestJson(req, 2 * 1024 * 1024);
    const date = String(body?.date || "").slice(0, 10);
    const entries = Array.isArray(body?.entries) ? body.entries : [];
    if (!date) {
      sendJson(res, 400, { error: "date is required" });
      return;
    }

    try {
      const settings = await readFustSettings();
      const syncedEntries = await syncDagFoutjesEntriesToSheet(entries, date, settings);
      sendJson(res, 200, { ok: true, date, entries: syncedEntries });
    } catch (error) {
      sendJson(res, 400, { error: error instanceof Error ? error.message : String(error) });
    }
    return;
  }

  if (url.pathname === "/api/dag-foutjes/sheet-delete" && req.method === "POST") {
    if (!requirePermission(res, requestUser, PERMISSIONS.EXPEDITION_STICKERS_VIEW)) {
      return;
    }

    const body = await readRequestJson(req, 512 * 1024);
    try {
      const settings = await readFustSettings();
      const payload = await clearDagFoutjesEntryFromSheet(body?.entry || {}, settings);
      sendJson(res, 200, payload);
    } catch (error) {
      sendJson(res, 400, { error: error instanceof Error ? error.message : String(error) });
    }
    return;
  }

  if (url.pathname === "/api/hal-locations/load-sheet" && req.method === "POST") {
    if (!requirePermission(res, requestUser, PERMISSIONS.HAL_LOCATIONS_VIEW)) {
      return;
    }

    let session = null;
    try {
      const settings = await readFustSettings();
      const spreadsheetId = String(settings.hal_locations_spreadsheet_id || settings.spreadsheet_id || "").trim();
      const sheetName = String(settings.hal_locations_sheet_name || "ERP_PASTE").trim() || "ERP_PASTE";
      if (!spreadsheetId) {
        sendJson(res, 400, { error: "Set a Hal Locations spreadsheet ID in Settings first" });
        return;
      }

      const rows = await loadSheetRows(spreadsheetId, sheetName);
      session = await createHalLocationSheetSession(rows, { spreadsheet_id: spreadsheetId, sheet_name: sheetName });
      const output = await runHalLocationsWorker([
        "inspect",
        "--input",
        session.file_path,
      ]);
      const payload = JSON.parse(output.toString("utf8"));
      sendJson(res, 200, {
        id: session.id,
        locPrefixes: Array.isArray(payload.locPrefixes) ? payload.locPrefixes : [],
        custPrefixes: Array.isArray(payload.custPrefixes) ? payload.custPrefixes : [],
        custByLoc: payload.custByLoc && typeof payload.custByLoc === "object" ? payload.custByLoc : {},
        totalRows: Number(payload.totalRows || 0),
        source: {
          type: "sheet",
          spreadsheet_id: spreadsheetId,
          sheet_name: sheetName,
        },
      });
    } catch (error) {
      if (session?.id) {
        halLocationSessions.delete(session.id);
      }
      if (session?.dir_path) {
        await fs.rm(session.dir_path, { recursive: true, force: true }).catch(() => {});
      }
      sendJson(res, 400, { error: error instanceof Error ? error.message : String(error) });
    }
    return;
  }

  if (url.pathname === "/api/hal-locations/generate" && req.method === "POST") {
    if (!requirePermission(res, requestUser, PERMISSIONS.HAL_LOCATIONS_VIEW)) {
      return;
    }

    const body = await readRequestJson(req, 2 * 1024 * 1024);
    const session = await getHalLocationSession(body.id);
    if (!session) {
      sendJson(res, 404, { error: "Session expired. Upload the halindeling again." });
      return;
    }

    const outputPath = path.join(session.dir_path, `stickers-${Date.now()}.pdf`);
    try {
      await runHalLocationsWorker([
        "generate",
        "--input",
        session.file_path,
        "--output",
        outputPath,
        "--loc-prefixes-json",
        JSON.stringify(normalizeHalPrefixList(body.locPrefixes)),
        "--cust-prefixes-json",
        JSON.stringify(normalizeHalPrefixList(body.custPrefixes)),
      ]);

      const pdfBuffer = await fs.readFile(outputPath);
      res.writeHead(200, {
        "content-type": "application/pdf",
        "content-disposition": `attachment; filename="stickers-${localDateIso()}.pdf"`,
        "cache-control": "private, no-store",
      });
      res.end(pdfBuffer);
    } catch (error) {
      sendJson(res, 400, { error: error instanceof Error ? error.message : String(error) });
    } finally {
      await fs.rm(outputPath, { force: true }).catch(() => {});
    }
    return;
  }

  if (url.pathname === "/api/cmrprint/data") {
    if (!requirePermission(res, requestUser, PERMISSIONS.CMR_VIEW)) {
      return;
    }
    try {
      const [payload, settings] = await Promise.all([loadCmrPrintData(), readFustSettings()]);
      sendJson(res, 200, {
        ...payload,
        settings: {
          cmr_default_template_name: settings.cmr_default_template_name,
          cmr_manage_usernames: settings.cmr_manage_usernames,
        },
        can_manage: canManageCmrWorkspace(requestUser, settings),
      });
    } catch (error) {
      sendJson(res, 500, { error: error instanceof Error ? error.message : String(error) });
    }
    return;
  }

  if (url.pathname === "/api/cmrprint/app-data" && req.method === "PATCH") {
    if (!requirePermission(res, requestUser, PERMISSIONS.CMR_VIEW)) {
      return;
    }
    const settings = await readFustSettings();
    if (!canManageCmrWorkspace(requestUser, settings)) {
      sendForbidden(res);
      return;
    }
    const body = await readRequestJson(req);
    await saveCmrPrintAppData(body);
    sendJson(res, 200, await loadCmrPrintData());
    return;
  }

  if (url.pathname === "/api/cmrprint/template" && req.method === "PUT") {
    if (!requirePermission(res, requestUser, PERMISSIONS.CMR_VIEW)) {
      return;
    }
    const settings = await readFustSettings();
    if (!canManageCmrWorkspace(requestUser, settings)) {
      sendForbidden(res);
      return;
    }
    const body = await readRequestJson(req);
    await saveCmrPrintTemplate(body.template || body);
    sendJson(res, 200, await loadCmrPrintData());
    return;
  }

  if (url.pathname === "/api/cmrprint/settings" && req.method === "PATCH") {
    if (!requirePermission(res, requestUser, PERMISSIONS.CMR_VIEW)) {
      return;
    }
    const settings = await readFustSettings();
    if (!canManageCmrWorkspace(requestUser, settings)) {
      sendForbidden(res);
      return;
    }
    const body = await readRequestJson(req);
    const nextSettings = normalizeFustSettings({
      ...settings,
      cmr_default_template_name: body?.cmr_default_template_name,
    });
    await writeFustSettings(nextSettings);
    sendJson(res, 200, {
      settings: {
        cmr_default_template_name: nextSettings.cmr_default_template_name,
        cmr_manage_usernames: nextSettings.cmr_manage_usernames,
      },
    });
    return;
  }

  if (url.pathname.startsWith("/api/cmrprint/template/") && req.method === "DELETE") {
    if (!requirePermission(res, requestUser, PERMISSIONS.CMR_VIEW)) {
      return;
    }
    const settings = await readFustSettings();
    if (!canManageCmrWorkspace(requestUser, settings)) {
      sendForbidden(res);
      return;
    }
    const templateName = decodeURIComponent(url.pathname.slice("/api/cmrprint/template/".length));
    await deleteCmrPrintTemplate(templateName);
    sendJson(res, 200, await loadCmrPrintData());
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
      const deletedRecord = records[recordIndex];
      records.splice(recordIndex, 1);
      await writeClockRecords(records);

      const settings = await readFustSettings();
      try {
        await rewriteClockBackupDates([deletedRecord.action_date], records, settings);
        await writeClockRecords(records);
      } catch (sheetError) {
        // Keep the app data updated even if the backup tab rewrite fails.
      }

      sendJson(res, 200, { ok: true, deleted_record_id: recordId });
      return;
    }

    if (req.method === "PATCH") {
      const body = await readRequestJson(req);
      const current = records[recordIndex];
      const previousDate = current.action_date;
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

      const settings = await readFustSettings();
      try {
        await rewriteClockBackupDates([previousDate, updated.action_date], records, settings);
        await writeClockRecords(records);
      } catch (sheetError) {
        // Keep the app data updated even if the backup tab rewrite fails.
      }

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

  if (url.pathname === "/api/ukdocs/state") {
    if (!requireAnyPermission(res, requestUser, [PERMISSIONS.UKDOCS_VIEW, PERMISSIONS.UKDOCS_INSPECTION_VIEW, PERMISSIONS.UKDOCS_CSI_VIEW])) {
      return;
    }

    if (req.method === "GET") {
      const state = await readUkdocsState();
      const settings = await readFustSettings();
      sendJson(res, 200, { state, settings: ukdocsViewerSettingsSummary(settings) });
      return;
    }

    if (req.method === "PATCH") {
      const body = await readRequestJson(req);
      const currentState = await readUkdocsState();
      const nextState = mergeUkdocsStatePatch(currentState, body);
      await writeUkdocsState(nextState);
      sendJson(res, 200, { state: nextState });
      return;
    }
  }

  if (url.pathname === "/api/ukdocs/import-examples" && req.method === "POST") {
    if (!requirePermission(res, requestUser, PERMISSIONS.UKDOCS_VIEW)) {
      return;
    }
    const body = await readRequestJson(req, 40 * 1024 * 1024);
    const output = await runUkdocsWorker(["import-examples"], JSON.stringify(body));
    sendJson(res, 200, JSON.parse(output.toString("utf8")));
    return;
  }

  if (url.pathname === "/api/ukdocs/analyze" && req.method === "POST") {
    if (!requirePermission(res, requestUser, PERMISSIONS.UKDOCS_VIEW)) {
      return;
    }
    const body = await readRequestJson(req, 60 * 1024 * 1024);
    const output = await runUkdocsWorker(["analyze"], JSON.stringify(body));
    sendJson(res, 200, JSON.parse(output.toString("utf8")));
    return;
  }

  if (url.pathname === "/api/ukdocs/generate" && req.method === "POST") {
    if (!requirePermission(res, requestUser, PERMISSIONS.UKDOCS_VIEW)) {
      return;
    }
    const body = await readRequestJson(req, 60 * 1024 * 1024);
    const generatedPayload = JSON.parse((await runUkdocsWorker(["generate"], JSON.stringify(body))).toString("utf8"));
    const analysis = generatedPayload.analysis || {};
    const state = await readUkdocsState();
    const warningMessages = Array.isArray(analysis?.audit?.warnings)
      ? analysis.audit.warnings.map((warning) => String(warning?.message || warning || "").trim()).filter(Boolean)
      : [];
    const shipment = normalizeUkdocsShipment({
      ...body,
      id: body?.id || crypto.randomUUID(),
      invoice_numbers: body?.invoice_numbers || "",
      uploaded_files: ukdocsUploadedFilesWithoutContent(body?.uploaded_files),
      validation_warnings: warningMessages,
      audit_status: analysis?.audit?.final_status === "PASS" ? "passed" : "failed",
      ready: analysis?.audit?.final_status === "PASS",
      created_by: body?.created_by || requestUser.username,
      created_at: body?.created_at || new Date().toISOString(),
      updated_at: new Date().toISOString(),
    });
    const existingIndex = state.shipments.findIndex((item) => item.id === shipment.id);
    if (existingIndex >= 0) {
      shipment.created_at = state.shipments[existingIndex].created_at || shipment.created_at;
      shipment.created_by = state.shipments[existingIndex].created_by || shipment.created_by;
      state.shipments[existingIndex] = shipment;
    } else {
      state.shipments.unshift(shipment);
    }

    const customer = state.customers.find((item) => item.id === shipment.customer_id);
    const existingCollection = findMatchingUkdocsPrintCollection(state.print_collections, {
      id: shipment.print_collection_id || shipment.id,
      shipment_id: shipment.id,
      shipment_date: shipment.shipment_date,
      reference_connect: shipment.reference_connect,
      invoice_numbers: shipment.invoice_numbers,
      truck_number: shipment.truck_number,
      trailer_number: shipment.trailer_number,
    });
    let printCollection = buildUkdocsPrintCollectionFromShipment(existingCollection, shipment, customer?.customer_name || "");
    const generatedFilesForCollection = await saveUkdocsGeneratedFiles(printCollection, (generatedPayload.files || []).filter((file) => file.kind !== "audit"), requestUser);
    printCollection = normalizeUkdocsPrintCollection({
      ...printCollection,
      documents: {
        ...(printCollection.documents || {}),
        generated_files: generatedFilesForCollection,
      },
    });
    state.print_collections = upsertUkdocsPrintCollection(state.print_collections, printCollection);
    const auditReport = normalizeUkdocsAuditReport({
      shipment_id: shipment.id,
      shipment_reference: shipment.export_reference || shipment.invoice_numbers || analysis?.shipment?.reference_line || "",
      shipment_date: shipment.shipment_date,
      customer_name: customer?.customer_name || "",
      created_at: new Date().toISOString(),
      final_status: analysis?.audit?.final_status || "",
      warnings: warningMessages,
      summary: summarizeUkdocsWarnings(analysis?.audit?.warnings) || `${analysis?.categories?.length || 0} categories checked`,
      summary_rows: analysis?.audit?.summary_rows || [],
    });
    state.audit_reports = [auditReport, ...state.audit_reports];
    await writeUkdocsState(state);
    const pdfJobs = await queueUkdocsInvoicePdfJobs(printCollection, requestUser);

    sendJson(res, 200, {
      analysis,
      files: (generatedPayload.files || []).filter((file) => file.kind !== "audit"),
      pdf_jobs: pdfJobs,
      shipment,
      shipments: normalizeUkdocsState(state).shipments,
      audit_reports: normalizeUkdocsState(state).audit_reports,
      print_collections: normalizeUkdocsState(state).print_collections,
    });
    return;
  }

  if (url.pathname === "/api/ukdocs/shipments") {
    if (!requirePermission(res, requestUser, PERMISSIONS.UKDOCS_VIEW)) {
      return;
    }

    if (req.method === "POST") {
      const body = await readRequestJson(req);
      const state = await readUkdocsState();
      const shipment = normalizeUkdocsShipment({
        ...body,
        created_by: body?.created_by || requestUser.username,
        created_at: body?.created_at || new Date().toISOString(),
        updated_at: new Date().toISOString(),
      });
      const existingIndex = state.shipments.findIndex((item) => item.id === shipment.id);
      if (existingIndex >= 0) {
        state.shipments[existingIndex] = shipment;
      } else {
        state.shipments.unshift(shipment);
      }
      await writeUkdocsState(state);
      sendJson(res, existingIndex >= 0 ? 200 : 201, { shipment, shipments: normalizeUkdocsState(state).shipments });
      return;
    }
  }

  if (url.pathname.startsWith("/api/ukdocs/shipments/") && req.method === "DELETE") {
    if (!requirePermission(res, requestUser, PERMISSIONS.UKDOCS_VIEW)) {
      return;
    }
    const shipmentId = decodeURIComponent(url.pathname.slice("/api/ukdocs/shipments/".length));
    const state = await readUkdocsState();
    const deletedShipment = state.shipments.find((item) => item.id === shipmentId) || null;
    const nextShipments = state.shipments.filter((item) => item.id !== shipmentId);
    if (nextShipments.length === state.shipments.length) {
      sendJson(res, 404, { error: "Shipment not found" });
      return;
    }
    state.shipments = nextShipments;
    if (deletedShipment?.print_collection_id) {
      const linkedCollection = state.print_collections.find((item) => item.id === deletedShipment.print_collection_id || item.shipment_id === deletedShipment.id);
      if (linkedCollection) {
        await deleteUkdocsGeneratedFiles(linkedCollection);
        const resetCollection = normalizeUkdocsPrintCollection({
          ...linkedCollection,
          shipment_id: "",
          shipment_reference: "",
          generated_at: "",
          updated_at: new Date().toISOString(),
          delivery_email: {
            ok: false,
            recipients: [],
            sent_at: "",
            error: "",
          },
          documents: {
            ...(linkedCollection.documents || {}),
            generated_files: [],
          },
        });
        state.print_collections = upsertUkdocsPrintCollection(state.print_collections, resetCollection);
      }
    }
    await writeUkdocsState(state);
    sendJson(res, 200, { ok: true, shipments: normalizeUkdocsState(state).shipments, print_collections: normalizeUkdocsState(state).print_collections });
    return;
  }

  if (url.pathname.startsWith("/api/ukdocs-print/collections/") && req.method === "PATCH") {
    if (!requireAnyPermission(res, requestUser, [PERMISSIONS.UKDOCS_VIEW, PERMISSIONS.UKDOCS_INSPECTION_VIEW, PERMISSIONS.UKDOCS_CSI_VIEW])) {
      return;
    }
    const collectionId = decodeURIComponent(url.pathname.slice("/api/ukdocs-print/collections/".length));
    const body = await readRequestJson(req);
    const state = await readUkdocsState();
    const existingCollection = state.print_collections.find((item) => item.id === collectionId || item.shipment_id === collectionId);
    if (!existingCollection) {
      sendJson(res, 404, { error: "UKdocs Print collection not found" });
      return;
    }
    const updatedCollection = normalizeUkdocsPrintCollection({
      ...existingCollection,
      notes: body?.notes ?? existingCollection.notes,
      updated_at: new Date().toISOString(),
    });
    state.print_collections = upsertUkdocsPrintCollection(state.print_collections, updatedCollection);
    await writeUkdocsState(state);
    sendJson(res, 200, { collection: updatedCollection, print_collections: normalizeUkdocsState(state).print_collections });
    return;
  }

  if (url.pathname.startsWith("/api/ukdocs-print/collections/") && url.pathname.endsWith("/send-ready") && req.method === "POST") {
    if (!requirePermission(res, requestUser, PERMISSIONS.UKDOCS_VIEW)) {
      return;
    }
    const collectionId = decodeURIComponent(url.pathname.slice("/api/ukdocs-print/collections/".length, -"/send-ready".length));
    const state = await readUkdocsState();
    const settings = await readFustSettings();
    const existingCollection = state.print_collections.find((item) => item.id === collectionId || item.shipment_id === collectionId);
    if (!existingCollection) {
      sendJson(res, 404, { error: "UKdocs Print collection not found" });
      return;
    }
    if (ukdocsPrintInspectionMode(existingCollection) === "stock_control") {
      sendJson(res, 400, { error: "Stock control collections do not send export papers" });
      return;
    }
    const deliveryEmail = await sendUkdocsPrintReadyEmail(existingCollection, state.customers, settings);
    const updatedCollection = normalizeUkdocsPrintCollection({
      ...existingCollection,
      updated_at: new Date().toISOString(),
      delivery_email: deliveryEmail,
    });
    state.print_collections = upsertUkdocsPrintCollection(state.print_collections, updatedCollection);
    await writeUkdocsState(state);
    sendJson(res, 200, { collection: updatedCollection, print_collections: normalizeUkdocsState(state).print_collections, delivery_email: deliveryEmail });
    return;
  }

  if (url.pathname.startsWith("/api/ukdocs-print/collections/") && !url.pathname.includes("/documents/") && req.method === "DELETE") {
    if (!requireAnyPermission(res, requestUser, [PERMISSIONS.UKDOCS_VIEW, PERMISSIONS.UKDOCS_INSPECTION_VIEW, PERMISSIONS.UKDOCS_CSI_VIEW])) {
      return;
    }
    const collectionId = decodeURIComponent(url.pathname.slice("/api/ukdocs-print/collections/".length));
    const state = await readUkdocsState();
    const existingCollection = state.print_collections.find((item) => item.id === collectionId || item.shipment_id === collectionId);
    if (!existingCollection) {
      sendJson(res, 404, { error: "UKdocs Print collection not found" });
      return;
    }
    await deleteUkdocsPrintCollectionFiles(existingCollection);
    state.print_collections = state.print_collections.filter((item) => item.id !== existingCollection.id && item.shipment_id !== existingCollection.shipment_id);
    await writeUkdocsState(state);
    sendJson(res, 200, { ok: true, print_collections: normalizeUkdocsState(state).print_collections });
    return;
  }

  if (url.pathname.startsWith("/api/ukdocs-print/collections/") && url.pathname.endsWith("/upload") && req.method === "POST") {
    if (!requireAnyPermission(res, requestUser, [PERMISSIONS.UKDOCS_VIEW, PERMISSIONS.UKDOCS_INSPECTION_VIEW, PERMISSIONS.UKDOCS_CSI_VIEW])) {
      return;
    }
    const basePath = url.pathname.slice("/api/ukdocs-print/collections/".length, -"/upload".length);
    const collectionId = decodeURIComponent(basePath);
    const body = await readRequestJson(req, 40 * 1024 * 1024);
    const kind = String(body?.kind || "").trim();
    const state = await readUkdocsState();
    const existingCollection = state.print_collections.find((item) => item.id === collectionId || item.shipment_id === collectionId);
    if (!existingCollection) {
      sendJson(res, 404, { error: "UKdocs Print collection not found" });
      return;
    }
    const originalName = path.basename(String(body?.file?.file_name || body?.file?.name || "").trim());
    if (kind === "phyto" && hasUkdocsPrintDocumentWithName(existingCollection.documents?.phyto_files, originalName)) {
      sendJson(res, 200, { collection: existingCollection, print_collections: normalizeUkdocsState(state).print_collections, skipped: true, reason: "phyto already exists" });
      return;
    }
    if (kind === "export_extra" && existingCollection.documents?.export_extra?.storage_name) {
      sendJson(res, 200, { collection: existingCollection, print_collections: normalizeUkdocsState(state).print_collections, skipped: true, reason: "export_extra already exists" });
      return;
    }
    if (kind === "temp_phyto" && hasUkdocsPrintDocumentWithName(existingCollection.documents?.temp_phyto_files, originalName)) {
      sendJson(res, 200, { collection: existingCollection, print_collections: normalizeUkdocsState(state).print_collections, skipped: true, reason: "temp_phyto already exists" });
      return;
    }
    if (["inspection_list", "locations_file", "export_extra", "temp_phyto_plants_file", "temp_phyto_plants_xml_file", "ipaffs_file", "ipaffs_plants_file"].includes(kind) && existingCollection.documents?.[kind]?.storage_name) {
      await deleteSingleUkdocsPrintDocumentFile(existingCollection.documents[kind]);
    }
    const savedDocumentRaw = await saveUkdocsPrintUpload(existingCollection.id, kind, body?.file || {}, requestUser);
    const savedDocument = ["temp_phyto", "temp_phyto_plants_file", "temp_phyto_plants_xml_file", "ipaffs_file", "ipaffs_plants_file"].includes(kind)
      ? await enrichUkdocsCsiStoredDocument(savedDocumentRaw, kind)
      : savedDocumentRaw;
    const updatedCollection = normalizeUkdocsPrintCollection({
      ...existingCollection,
      updated_at: new Date().toISOString(),
      documents: {
        ...(existingCollection.documents || {}),
        ...(kind === "phyto"
          ? { phyto_files: [...(existingCollection.documents?.phyto_files || []), savedDocument] }
          : kind === "temp_phyto"
            ? { temp_phyto_files: [...(existingCollection.documents?.temp_phyto_files || []), savedDocument] }
            : { [kind]: savedDocument }),
      },
      csi_report: shouldResetUkdocsCsiReportForDocumentKind(kind)
        ? createEmptyUkdocsCsiReport()
        : existingCollection.csi_report,
    });
    state.print_collections = upsertUkdocsPrintCollection(state.print_collections, updatedCollection);
    await writeUkdocsState(state);
    sendJson(res, 200, { collection: updatedCollection, print_collections: normalizeUkdocsState(state).print_collections });
    return;
  }

  if (url.pathname.startsWith("/api/ukdocs-print/collections/") && url.pathname.includes("/documents/") && req.method === "GET") {
    if (!requireAnyPermission(res, requestUser, [PERMISSIONS.UKDOCS_VIEW, PERMISSIONS.UKDOCS_INSPECTION_VIEW, PERMISSIONS.UKDOCS_CSI_VIEW])) {
      return;
    }
    const suffix = url.pathname.slice("/api/ukdocs-print/collections/".length);
    const [collectionIdRaw, kindRaw] = suffix.split("/documents/");
    const collectionId = decodeURIComponent(collectionIdRaw || "");
    const kindParts = decodeURIComponent(kindRaw || "").split("/");
    const kind = kindParts[0] || "";
    const documentIndex = Number(kindParts[1] || 0);
    const state = await readUkdocsState();
    const collection = state.print_collections.find((item) => item.id === collectionId || item.shipment_id === collectionId);
    const document = kind === "phyto"
      ? (collection?.documents?.phyto_files || [])[documentIndex]
      : kind === "temp_phyto"
        ? (collection?.documents?.temp_phyto_files || [])[documentIndex]
      : kind === "generated"
        ? (collection?.documents?.generated_files || [])[documentIndex]
        : collection?.documents?.[kind];
    if (!collection || !document?.storage_name) {
      sendText(res, 404, "UKdocs Print document not found");
      return;
    }
    const resolvedPath = path.resolve(ukdocsPrintDocumentPath(document));
    if (!resolvedPath.startsWith(path.resolve(ukdocsPrintFilesDir)) || !existsSync(resolvedPath)) {
      sendText(res, 404, "Stored document not found");
      return;
    }
    res.writeHead(200, {
      "content-type": document.mime_type || guessMimeType(document.original_name),
      "content-disposition": `attachment; filename="${path.basename(document.original_name || resolvedPath)}"`,
      "cache-control": "no-store",
    });
    createReadStream(resolvedPath).pipe(res);
    return;
  }

  if (url.pathname.startsWith("/api/ukdocs-print/collections/") && url.pathname.includes("/documents/") && req.method === "DELETE") {
    if (!requireAnyPermission(res, requestUser, [PERMISSIONS.UKDOCS_VIEW, PERMISSIONS.UKDOCS_INSPECTION_VIEW, PERMISSIONS.UKDOCS_CSI_VIEW])) {
      return;
    }
    const suffix = url.pathname.slice("/api/ukdocs-print/collections/".length);
    const [collectionIdRaw, kindRaw] = suffix.split("/documents/");
    const collectionId = decodeURIComponent(collectionIdRaw || "");
    const kindParts = decodeURIComponent(kindRaw || "").split("/");
    const kind = kindParts[0] || "";
    const documentIndex = Number(kindParts[1] || 0);
    const state = await readUkdocsState();
    const existingCollection = state.print_collections.find((item) => item.id === collectionId || item.shipment_id === collectionId);
    if (!existingCollection) {
      sendJson(res, 404, { error: "UKdocs Print collection not found" });
      return;
    }

    let removedDocument = null;
    let updatedDocuments = { ...(existingCollection.documents || {}) };

    if (kind === "phyto") {
      const phytoFiles = [...(existingCollection.documents?.phyto_files || [])];
      removedDocument = phytoFiles[documentIndex] || null;
      if (!removedDocument) {
        sendJson(res, 404, { error: "UKdocs Print document not found" });
        return;
      }
      phytoFiles.splice(documentIndex, 1);
      updatedDocuments = { ...updatedDocuments, phyto_files: phytoFiles };
    } else if (kind === "temp_phyto") {
      const tempPhytoFiles = [...(existingCollection.documents?.temp_phyto_files || [])];
      removedDocument = tempPhytoFiles[documentIndex] || null;
      if (!removedDocument) {
        sendJson(res, 404, { error: "UKdocs Print document not found" });
        return;
      }
      tempPhytoFiles.splice(documentIndex, 1);
      updatedDocuments = { ...updatedDocuments, temp_phyto_files: tempPhytoFiles };
    } else if (kind === "generated") {
      const generatedFiles = [...(existingCollection.documents?.generated_files || [])];
      removedDocument = generatedFiles[documentIndex] || null;
      if (!removedDocument) {
        sendJson(res, 404, { error: "UKdocs Print document not found" });
        return;
      }
      generatedFiles.splice(documentIndex, 1);
      updatedDocuments = { ...updatedDocuments, generated_files: generatedFiles };
    } else if (["export_extra", "inspection_list", "locations_file", "temp_phyto_plants_file", "temp_phyto_plants_xml_file", "ipaffs_file", "ipaffs_plants_file"].includes(kind)) {
      removedDocument = existingCollection.documents?.[kind] || null;
      if (!removedDocument) {
        sendJson(res, 404, { error: "UKdocs Print document not found" });
        return;
      }
      updatedDocuments = { ...updatedDocuments, [kind]: null };
    } else {
      sendJson(res, 400, { error: "Unknown UKdocs Print document type" });
      return;
    }

    await deleteSingleUkdocsPrintDocumentFile(removedDocument);
    const updatedCollection = normalizeUkdocsPrintCollection({
      ...existingCollection,
      updated_at: new Date().toISOString(),
      documents: updatedDocuments,
      csi_report: shouldResetUkdocsCsiReportForDocumentKind(kind)
        ? createEmptyUkdocsCsiReport()
        : existingCollection.csi_report,
    });
    state.print_collections = upsertUkdocsPrintCollection(state.print_collections, updatedCollection);
    await writeUkdocsState(state);
    sendJson(res, 200, { collection: updatedCollection, print_collections: normalizeUkdocsState(state).print_collections });
    return;
  }

  if (url.pathname.startsWith("/api/ukdocs-print/collections/") && url.pathname.endsWith("/csi/run") && req.method === "POST") {
    if (!requirePermission(res, requestUser, PERMISSIONS.UKDOCS_CSI_VIEW)) {
      return;
    }
    const collectionId = decodeURIComponent(url.pathname.slice("/api/ukdocs-print/collections/".length, -"/csi/run".length));
    const state = await readUkdocsState();
    const existingCollection = ukdocsPrintCollectionById(state.print_collections, collectionId);
    if (!existingCollection) {
      sendJson(res, 404, { error: "UKDocs zending not found" });
      return;
    }
    const generatedFiles = existingCollection.documents?.generated_files || [];
    const hasGeneratedExport = generatedFiles.some((file) => file.document_kind === "export");
    const hasGeneratedInvoices = generatedFiles.some((file) => file.document_kind === "invoice");
    if (!hasGeneratedExport || !hasGeneratedInvoices) {
      sendJson(res, 400, { error: "Generate the export and invoice files before running CSI" });
      return;
    }
    if (
      !isUkdocsNoPdNeeded(existingCollection)
      && !existingCollection.documents?.ipaffs_file?.storage_name
      && !existingCollection.documents?.ipaffs_plants_file?.storage_name
    ) {
      sendJson(res, 400, { error: "Upload the IPAFFS file before running CSI" });
      return;
    }
    const hydrated = await hydrateUkdocsCsiCollectionInputs(existingCollection, { force_refresh: true });
    let collectionForRun = hydrated.collection || existingCollection;
    if (hydrated.changed && collectionForRun) {
      state.print_collections = upsertUkdocsPrintCollection(state.print_collections, collectionForRun);
      await writeUkdocsState(state);
    }
    if (isUkdocsNoPdNeeded(collectionForRun)) {
      const extractedDocuments = await extractUkdocsCsiFileSnapshots([
        ...(collectionForRun?.documents?.generated_files || []).map((document) => ({ kind: document.document_kind === "export" ? "generated_export" : "generated_invoice", document })),
      ]);
      const deterministicBundle = buildUkdocsCsiDeterministicReport(collectionForRun, extractedDocuments);
      const completedCollection = await updateUkdocsCsiReport(collectionForRun.id, {
        status: "done",
        job_id: "",
        queued_at: new Date().toISOString(),
        started_at: new Date().toISOString(),
        completed_at: new Date().toISOString(),
        error: "",
        summary: deterministicBundle.report.summary || "CSI audit completed.",
        overall_status: deterministicBundle.report.overall_status || "pass",
        checks: deterministicBundle.report.checks || [],
        products: deterministicBundle.report.products || [],
        flower_products: deterministicBundle.report.flower_products || [],
        plant_products: deterministicBundle.report.plant_products || [],
        manual_checks: deterministicBundle.report.manual_checks || [],
        notes: uniqueUkdocsCsiStrings([
          ...(deterministicBundle.report.notes || []),
          "PD code indicates no PD needed, so CSI passed from generated invoice/export checks.",
        ]),
        llm_content: "",
        llm_parse_source: "",
        llm_parse_error: "",
        llm_raw_result_json: "",
      });
      await updateUkdocsCsiEmailResult(collectionForRun.id, {
        ok: false,
        recipients: [],
        sent_at: "",
        error: "",
      });
      const nextState = await readUkdocsState();
      sendJson(res, 200, {
        ok: true,
        collection: completedCollection,
        print_collections: normalizeUkdocsState(nextState).print_collections,
      });
      return;
    }
    const queuedJobs = await queueUkdocsCsiAudit(collectionForRun, requestUser);
    const nextState = await readUkdocsState();
    sendJson(res, 200, {
      ok: true,
      job: Array.isArray(queuedJobs) ? (queuedJobs[0] || null) : queuedJobs,
      jobs: Array.isArray(queuedJobs) ? queuedJobs : [queuedJobs].filter(Boolean),
      collection: ukdocsPrintCollectionById(nextState.print_collections, existingCollection.id),
      print_collections: normalizeUkdocsState(nextState).print_collections,
    });
    return;
  }

  if (url.pathname.startsWith("/api/ukdocs-print/collections/") && url.pathname.endsWith("/csi/send") && req.method === "POST") {
    if (!requirePermission(res, requestUser, PERMISSIONS.UKDOCS_CSI_VIEW)) {
      return;
    }
    const collectionId = decodeURIComponent(url.pathname.slice("/api/ukdocs-print/collections/".length, -"/csi/send".length));
    const state = await readUkdocsState();
    const existingCollection = ukdocsPrintCollectionById(state.print_collections, collectionId);
    if (!existingCollection) {
      sendJson(res, 404, { error: "UKDocs zending not found" });
      return;
    }
    const overallStatus = String(existingCollection?.csi_report?.overall_status || "").trim().toLowerCase();
    if (String(existingCollection?.csi_report?.status || "").trim() !== "done" || overallStatus !== "pass") {
      sendJson(res, 400, { error: "Run CSI successfully before sending papers to CSI" });
      return;
    }
    const settings = await readFustSettings();
    const csiEmail = await sendUkdocsCsiSuccessEmail(existingCollection, state.customers, settings);
    const updatedCollection = await updateUkdocsCsiEmailResult(collectionId, csiEmail);
    const nextState = await readUkdocsState();
    sendJson(res, 200, {
      ok: true,
      collection: updatedCollection,
      print_collections: normalizeUkdocsState(nextState).print_collections,
      csi_email: csiEmail,
    });
    return;
  }

  if (url.pathname === "/api/fust/settings") {
    if (!requirePermission(res, requestUser, PERMISSIONS.SETTINGS_MANAGE)) {
      return;
    }

    if (req.method === "GET") {
      const settings = await readFustSettings();
      const cmrData = await loadCmrPrintData();
      sendJson(res, 200, { settings: {
        ...settings,
        cmr_data_dir: cmrData.data_dir,
        cmr_templates_dir: cmrData.templates_dir,
        cmr_available_templates: (cmrData.templates || []).map((item) => item.name),
      } });
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
      const cmrData = await loadCmrPrintData();
      sendJson(res, 200, { settings: {
        ...nextSettings,
        cmr_data_dir: cmrData.data_dir,
        cmr_templates_dir: cmrData.templates_dir,
        cmr_available_templates: (cmrData.templates || []).map((item) => item.name),
      } });
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
      const tokenPayload = await exchangeGoogleAuthCode(settings, code, cmrGoogleRedirectUri(req));
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

  if (url.pathname === "/api/ukdocs-print/gmail/auth-url") {
    if (!requirePermission(res, requestUser, PERMISSIONS.SETTINGS_MANAGE)) {
      return;
    }
    const settings = await readFustSettings();
    if (!settings.cmr_google_client_id || !settings.cmr_google_client_secret) {
      sendJson(res, 400, { error: "Set Google OAuth client ID and secret first" });
      return;
    }
    sendJson(res, 200, { auth_url: ukdocsGmailAuthUrl(settings, req), redirect_uri: ukdocsGmailRedirectUri(req) });
    return;
  }

  if (url.pathname === "/api/ukdocs-print/gmail/callback") {
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
      const tokenPayload = await exchangeGoogleAuthCode(settings, code, ukdocsGmailRedirectUri(req));
      if (!tokenPayload.refresh_token) {
        sendText(res, 400, "Google did not return a refresh token. Try Connect Gmail again and approve offline access.");
        return;
      }
      const connectedEmail = await loadGoogleUserEmail(tokenPayload.access_token);
      await writeFustSettings({
        ...settings,
        gmail_refresh_token: tokenPayload.refresh_token,
        gmail_connected_email: connectedEmail,
      });
      res.writeHead(200, { "content-type": "text/html; charset=utf-8" });
      res.end("<p>Gmail connected for UKdocs Print. You can close this tab and return to SnappySjaak.</p>");
    } catch (error) {
      sendText(res, 500, error instanceof Error ? error.message : String(error));
    }
    return;
  }

  if (url.pathname === "/api/ukdocs-print/gmail/sync" && req.method === "POST") {
    if (!requirePermission(res, requestUser, PERMISSIONS.UKDOCS_VIEW)) {
      return;
    }
    const body = await readRequestJson(req);
    const settings = await readFustSettings();
    try {
      const payload = await syncUkdocsPrintFromGmail(settings, requestUser, body?.query, body?.date || localDateIso());
      sendJson(res, 200, payload);
    } catch (error) {
      const serviceLabel = "UKdocs Gmail";
      const reconnectTarget = settings.gmail_connected_email || "the Gmail account used for UKdocs Print";
      await sendSupportAttentionEmail(settings, {
        service: serviceLabel,
        action: "Sync Gmail attachments",
        username: requestUser.username,
        connected_account: settings.gmail_connected_email || "",
        reconnect_target: reconnectTarget,
        workaround: buildReconnectGuidance(serviceLabel, settings.gmail_connected_email, reconnectTarget),
        error: error instanceof Error ? error.message : String(error),
        path: url.pathname,
      });
      sendJson(res, 400, {
        error: `${error instanceof Error ? error.message : String(error)} ${buildReconnectGuidance(serviceLabel, settings.gmail_connected_email, reconnectTarget)}`,
      });
    }
    return;
  }

  if (url.pathname === "/api/ukdocs-print/sheet-sync" && req.method === "POST") {
    if (!requireAnyPermission(res, requestUser, [PERMISSIONS.UKDOCS_VIEW, PERMISSIONS.UKDOCS_INSPECTION_VIEW])) {
      return;
    }
    const body = await readRequestJson(req);
    const settings = await readFustSettings();
    try {
      const payload = await syncUkdocsPrintCollectionsFromSheet(settings, body?.date || localDateIso(), {
        reference_connect_only: body?.reference_connect_only === true,
        update_only: body?.update_only === true,
        overwrite_existing: body?.overwrite_existing === true,
      });
      sendJson(res, 200, payload);
    } catch (error) {
      let serviceAccount = "";
      try {
        serviceAccount = String((await loadServiceAccountInfo())?.client_email || "");
      } catch {
      }
      const serviceLabel = "UKdocs spreadsheet";
      const reconnectTarget = serviceAccount || "the Google Sheets service account";
      await sendSupportAttentionEmail(settings, {
        service: serviceLabel,
        action: body?.update_only === true ? "Update shipment info from spreadsheet" : "Load sendings from spreadsheet",
        username: requestUser.username,
        connected_account: serviceAccount,
        reconnect_target: reconnectTarget,
        workaround: `Check that the spreadsheet is shared with ${reconnectTarget} and that the service account credentials are still valid.`,
        error: error instanceof Error ? error.message : String(error),
        path: url.pathname,
      });
      sendJson(res, 400, {
        error: `${error instanceof Error ? error.message : String(error)} Check that the spreadsheet is shared with ${reconnectTarget} and that the service account credentials are still valid.`,
      });
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

  if (url.pathname === "/api/fust/backups/restore-missing" && req.method === "POST") {
    if (!requirePermission(res, requestUser, PERMISSIONS.SETTINGS_MANAGE)) {
      return;
    }
    const body = await readRequestJson(req);
    try {
      const payload = await restoreMissingFustDataFromBackup(body?.filename, requestUser);
      sendJson(res, 200, payload);
    } catch (error) {
      sendJson(res, 400, { error: error instanceof Error ? error.message : String(error) });
    }
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
    let exporters = [];
    let cmrCustomers = [];

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

    try {
      const cmrData = await loadCmrPrintData();
      exporters = (cmrData?.exporters || []).map(exporterBlockFromCmrProfile).filter((item) => item.name);
      cmrCustomers = (cmrData?.customers || []).map((item) => ({
        name: String(item?.name || "").trim(),
        exporter_profile_name: String(item?.exporter_profile_name || "").trim(),
      })).filter((item) => item.name);
    } catch {
      exporters = [];
      cmrCustomers = [];
    }

    const countries = [...new Set(records.map((record) => record.country))].sort((left, right) => left.localeCompare(right));
    sendJson(res, 200, {
      settings,
      countries,
      exporters,
      cmr_customers: cmrCustomers,
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
      database: getDatabaseStatus(),
      database_stats: getDatabaseStatus().ready ? await getFustDatabaseStats().catch(() => null) : null,
      error,
    });
    return;
  }

  if (url.pathname === "/api/fust/database/backfill" && req.method === "POST") {
    if (!requirePermission(res, requestUser, PERMISSIONS.SETTINGS_MANAGE)) {
      return;
    }
    try {
      const payload = await backfillFustDatabase();
      sendJson(res, 200, payload);
    } catch (error) {
      sendJson(res, 400, { error: error instanceof Error ? error.message : String(error) });
    }
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
      const settings = await readFustSettings();
      const localMatch = await ensureLocalFustAction(actionId, settings);
      const action = localMatch.action;
      const documentInfo = normalizeCmrInfo(action?.[documentKind]);
      if (!action || documentInfo.status !== "uploaded" || !documentInfo.file_id) {
        sendText(res, 404, "Document not found");
        return;
      }
      try {
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

    const { settings, localActions, actions, sourceDebug } = await loadCurrentFustActionsSnapshot();
    const country = String(url.searchParams.get("country") || "").trim();
    const customer = String(url.searchParams.get("customer_name") || "").trim().toLowerCase();
    const type = String(url.searchParams.get("type") || "").trim().toUpperCase();
    const filteredActions = actions
      .filter((action) => !country || action.country === country)
      .filter((action) => !customer || action.customer_name.toLowerCase().includes(customer))
      .filter((action) => !type || action.type === type);
    const reminderStatus = await maybeSendFustConfirmationReminders(actions, settings);

    sendJson(res, 200, {
      actions: filteredActions.sort((left, right) => {
        const rightDate = String(right.created_at || right.action_date || "");
        const leftDate = String(left.created_at || left.action_date || "");
        return rightDate.localeCompare(leftDate);
      }),
      overview: buildOverview(filteredActions),
      controle_summary: reminderStatus,
      source_debug: {
        ...sourceDebug,
        local: {
          action_count: localActions.filter((action) => !action.deleted).length,
          deleted_action_count: localActions.filter((action) => action.deleted).length,
        },
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

    const settings = await readFustSettings();
    const localMatch = await ensureLocalFustAction(actionId, settings);
    const actions = localMatch.actions;
    const actionIndex = localMatch.actionIndex;
    if (actionIndex < 0 || !localMatch.action) {
      sendJson(res, 404, { error: "Fust action not found" });
      return;
    }

    const action = localMatch.action;
    if (action.type !== documentConfig.type) {
      sendJson(res, 400, { error: `${documentConfig.label} files can only be attached to ${documentConfig.type} actions` });
      return;
    }
    const requiredPermission = action.type === "OUT" ? PERMISSIONS.FUST_OUT : PERMISSIONS.FUST_IN;
    if (!requirePermission(res, requestUser, requiredPermission)) {
      return;
    }
    if (isFustActionConfirmed(action) && !hasUserPermission(requestUser, PERMISSIONS.FUST_MANAGE)) {
      sendJson(res, 403, { error: "Confirmed actions can only be changed in Fust Beheer" });
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
      await mirrorFustActionToDatabase(action);
      sendJson(res, 200, { action });
      return;
    }

    const body = await readRequestJson(req, 18 * 1024 * 1024);
    const filePayload = body?.file || {};
    if (!filePayload.content_base64 || !filePayload.name) {
      sendJson(res, 400, { error: `Choose a ${documentConfig.label} file first` });
      return;
    }

    try {
      action[documentConfig.field] = normalizeCmrInfo({
        ...(await uploadFustDocumentToDrive(action, settings, filePayload, documentConfig.field)),
        uploaded_at: new Date().toISOString(),
        uploaded_by: requestUser.username,
      });
    } catch (documentError) {
      const serviceLabel = "Google Drive";
      const reconnectTarget = settings.cmr_google_connected_email || "the Google Drive upload account";
      await sendSupportAttentionEmail(settings, {
        service: serviceLabel,
        action: `Upload ${documentConfig.label}`,
        username: requestUser.username,
        connected_account: settings.cmr_google_connected_email || "",
        reconnect_target: reconnectTarget,
        workaround: buildReconnectGuidance(serviceLabel, settings.cmr_google_connected_email, reconnectTarget),
        error: documentError instanceof Error ? documentError.message : String(documentError),
        path: url.pathname,
      });
      action[documentConfig.field] = normalizeCmrInfo({
        status: "failed",
        error: `${documentError instanceof Error ? documentError.message : String(documentError)} ${buildReconnectGuidance(serviceLabel, settings.cmr_google_connected_email, reconnectTarget)}`,
        uploaded_at: new Date().toISOString(),
        uploaded_by: requestUser.username,
      });
      actions[actionIndex] = action;
      await writeFustActions(actions);
      await mirrorFustActionToDatabase(action);
      sendJson(res, 500, { error: action[documentConfig.field].error, action });
      return;
    }

    actions[actionIndex] = action;
    await writeFustActions(actions);
    await mirrorFustActionToDatabase(action);
    sendJson(res, 200, { action });
    return;
  }

  if (url.pathname.startsWith("/api/fust/actions/") && req.method === "POST") {
    const parts = url.pathname.split("/").filter(Boolean);
    const actionId = decodeURIComponent(parts[3] || "");
    const retryKind = parts[4] || "";
    if (actionId === "confirm-batch") {
      if (!requireAnyPermission(res, requestUser, [PERMISSIONS.FUST_OVERVIEW, PERMISSIONS.FUST_MANAGE])) {
        return;
      }
      const settings = await readFustSettings();
      const body = await readRequestJson(req);
      const actionIds = Array.isArray(body?.action_ids)
        ? body.action_ids.map((value) => String(value || "").trim()).filter(Boolean)
        : [];
      if (!actionIds.length) {
        sendJson(res, 400, { error: "No actions selected for confirmation" });
        return;
      }

      const results = {
        requested: actionIds.length,
        confirmed: 0,
        already_confirmed: 0,
        missing: 0,
      };

      for (const batchActionId of actionIds) {
        const localMatch = await ensureLocalFustAction(batchActionId, settings);
        if (localMatch.actionIndex < 0 || !localMatch.action) {
          results.missing += 1;
          continue;
        }
        const action = localMatch.action;
        if (isFustActionConfirmed(action)) {
          results.already_confirmed += 1;
          continue;
        }
        action.confirmed_at = new Date().toISOString();
        action.confirmed_by = requestUser.username;
        localMatch.actions[localMatch.actionIndex] = normalizeFustAction(action);
        await writeFustActions(localMatch.actions);
        await mirrorFustActionToDatabase(localMatch.actions[localMatch.actionIndex]);
        results.confirmed += 1;
      }

      sendJson(res, 200, { ok: true, summary: results });
      return;
    }
    const settings = await readFustSettings();
    const localMatch = await ensureLocalFustAction(actionId, settings);
    const actions = localMatch.actions;
    const actionIndex = localMatch.actionIndex;
    if (actionIndex < 0 || !localMatch.action) {
      sendJson(res, 404, { error: "Fust action not found" });
      return;
    }

    const action = localMatch.action;

    if (retryKind === "confirm" || retryKind === "unconfirm") {
      if (retryKind === "confirm") {
        if (!requireAnyPermission(res, requestUser, [PERMISSIONS.FUST_OVERVIEW, PERMISSIONS.FUST_MANAGE])) {
          return;
        }
        action.confirmed_at = new Date().toISOString();
        action.confirmed_by = requestUser.username;
      } else {
        if (!requirePermission(res, requestUser, PERMISSIONS.FUST_MANAGE)) {
          return;
        }
        action.confirmed_at = "";
        action.confirmed_by = "";
      }
      actions[actionIndex] = normalizeFustAction(action);
      await writeFustActions(actions);
      await mirrorFustActionToDatabase(actions[actionIndex]);
      sendJson(res, 200, { action: actions[actionIndex] });
      return;
    }

    if (retryKind === "retry-sheet") {
      const requiredPermission = action.type === "OUT" ? PERMISSIONS.FUST_OUT : PERMISSIONS.FUST_IN;
      if (!requirePermission(res, requestUser, requiredPermission)) {
        return;
      }
      if (isFustActionConfirmed(action) && !hasUserPermission(requestUser, PERMISSIONS.FUST_MANAGE)) {
        sendJson(res, 403, { error: "Confirmed actions can only be changed in Fust Beheer" });
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
      await mirrorFustActionToDatabase(action);
      sendJson(res, 200, { action });
      return;
    }

    if (retryKind === "retry-email") {
      const requiredPermission = action.type === "OUT" ? PERMISSIONS.FUST_OUT : PERMISSIONS.FUST_IN;
      if (!requirePermission(res, requestUser, requiredPermission)) {
        return;
      }
      if (isFustActionConfirmed(action) && !hasUserPermission(requestUser, PERMISSIONS.FUST_MANAGE)) {
        sendJson(res, 403, { error: "Confirmed actions can only be changed in Fust Beheer" });
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
      await mirrorFustActionToDatabase(action);
      sendJson(res, 200, { action });
      return;
    }

    sendJson(res, 404, { error: "Unknown retry action" });
    return;
  }

  if (url.pathname.startsWith("/api/fust/actions/") && req.method === "PUT") {
    const parts = url.pathname.split("/").filter(Boolean);
    const actionId = decodeURIComponent(parts[3] || "");
    const settings = await readFustSettings();
    const localMatch = await ensureLocalFustAction(actionId, settings);
    const actions = localMatch.actions;
    const actionIndex = localMatch.actionIndex;
    if (actionIndex < 0 || !localMatch.action) {
      sendJson(res, 404, { error: "Fust action not found" });
      return;
    }

    const existingAction = localMatch.action;
    const type = String(existingAction.type || "").trim().toUpperCase() === "OUT" ? "OUT" : "IN";
    const requiredPermission = type === "OUT" ? PERMISSIONS.FUST_OUT : PERMISSIONS.FUST_IN;
    if (!requirePermission(res, requestUser, requiredPermission)) {
      return;
    }
    if (isFustActionConfirmed(existingAction) && !hasUserPermission(requestUser, PERMISSIONS.FUST_MANAGE)) {
      sendJson(res, 403, { error: "Confirmed actions can only be changed in Fust Beheer" });
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
      fustbon_reference: body.fustbon_reference,
      fustfactuur_reference: body.fustfactuur_reference,
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
    await mirrorFustActionToDatabase(updatedAction);

    try {
      updatedAction.sheet_sync = await syncFustActionToSheets(updatedAction, settings, { previousAction: existingAction });
    } catch (sheetError) {
      updatedAction.sheet_sync = {
        ...(updatedAction.sheet_sync || {}),
        ok: false,
        target_sheets: [],
        error: sheetError instanceof Error ? sheetError.message : String(sheetError),
      };
    }

    actions[actionIndex] = updatedAction;
    await writeFustActions(actions);
    await mirrorFustActionToDatabase(updatedAction);
    sendJson(res, 200, { action: updatedAction });
    return;
  }

  if (url.pathname.startsWith("/api/fust/actions/") && req.method === "DELETE") {
    const parts = url.pathname.split("/").filter(Boolean);
    const actionId = decodeURIComponent(parts[3] || "");
    const actions = await readFustActions();
    let actionIndex = actions.findIndex((item) => item.id === actionId);
    let action = actionIndex >= 0 ? actions[actionIndex] : null;
    if (!action) {
      const settings = await readFustSettings();
      let candidates = [];
      try {
        const [retourRows, uitgaandRows] = await Promise.all([
          loadSheetRows(settings.spreadsheet_id, settings.in_sheet_name).catch(() => []),
          loadSheetRows(settings.spreadsheet_id, settings.out_sheet_name).catch(() => []),
        ]);
        candidates = [
          ...parseRegistrySheetRows(retourRows, "IN"),
          ...parseRegistrySheetRows(uitgaandRows, "OUT"),
        ];
      } catch {
        candidates = [];
      }
      action = candidates.find((item) => item.id === actionId) || null;
      if (!action) {
        sendJson(res, 404, { error: "Fust action not found" });
        return;
      }
      actionIndex = -1;
    }

    const requiredPermission = action.type === "OUT" ? PERMISSIONS.FUST_OUT : PERMISSIONS.FUST_IN;
    if (!requirePermission(res, requestUser, requiredPermission)) {
      return;
    }
    if (isFustActionConfirmed(action) && !hasUserPermission(requestUser, PERMISSIONS.FUST_MANAGE)) {
      sendJson(res, 403, { error: "Confirmed actions can only be changed in Fust Beheer" });
      return;
    }

    const settings = await readFustSettings();
    let deleteSync;
    try {
      deleteSync = await deleteFustActionFromSheets(action, settings);
    } catch (deleteError) {
      sendJson(res, 500, {
        error: deleteError instanceof Error ? deleteError.message : String(deleteError),
      });
      return;
    }

    let nextActions = actions;
    if (actionIndex >= 0) {
      nextActions = actions.filter((item) => item.id !== actionId);
      await writeFustActions(nextActions);
    }
    await mirrorFustDeleteToDatabase(actionId);

    sendJson(res, 200, {
      ok: true,
      deleted_action_id: actionId,
      sheet_sync: deleteSync,
    });
    return;
  }

  if (url.pathname === "/api/fust/fust-list" && req.method === "POST") {
    if (!requirePermission(res, requestUser, PERMISSIONS.FUST_IN)) {
      return;
    }

    const body = await readRequestJson(req, 512 * 1024);
    const actionDate = String(body.action_date || localDateIso()).trim();
    const customerName = String(body.customer_name || "").trim();
    const exporterInfo = normalizeFustExporterInfo(body.exporter || {});
    const rows = Array.isArray(body.rows) ? body.rows.map(normalizeFustListRow).filter((row) => row.code) : [];
    if (!customerName) {
      sendJson(res, 400, { error: "Customer is required" });
      return;
    }
    if (!rows.length) {
      sendJson(res, 400, { error: "Add at least one Fust Lijst row" });
      return;
    }

    try {
      const workbook = await generateFustListWorkbook({
        action_date: actionDate,
        customer_name: customerName,
        exporter: exporterInfo,
        rows,
        generated_by: requestUser.username,
      });
      const filename = contentDispositionFilename(`fust-lijst-${actionDate}-${customerName}.xlsx`);
      res.writeHead(200, {
        "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "content-disposition": `attachment; filename="${filename}"`,
        "cache-control": "no-store",
      });
      res.end(workbook);
    } catch (error) {
      sendJson(res, 400, { error: error instanceof Error ? error.message : String(error) });
    }
    return;
  }

  if (url.pathname === "/api/fust/submit" && req.method === "POST") {
    const body = await readRequestJson(req, 18 * 1024 * 1024);
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
      fustbon_reference: body.fustbon_reference,
      fustfactuur_reference: body.fustfactuur_reference,
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

    const documentPayload = body.document || {};
    const documentMode = String(documentPayload.mode || "").trim().toLowerCase();
    const documentLabel = type === "IN" ? "Fustbon" : "CMR";
    if (!documentMode) {
      sendJson(res, 400, { error: `Choose a ${documentLabel} file or mark No ${documentLabel}` });
      return;
    }
    if (!["upload", "skip"].includes(documentMode)) {
      sendJson(res, 400, { error: `Unknown ${documentLabel} choice` });
      return;
    }
    if (documentMode === "upload" && (!documentPayload.file?.content_base64 || !documentPayload.file?.name)) {
      sendJson(res, 400, { error: `Choose a ${documentLabel} file first` });
      return;
    }

    const actions = await readFustActions();
    actions.push(action);
    await writeFustActions(actions);
    await mirrorFustActionToDatabase(action);

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

    try {
      await applyFustDocumentChoice(action, settings, documentPayload, requestUser);
    } catch (documentChoiceError) {
      sendJson(res, 400, { error: documentChoiceError instanceof Error ? documentChoiceError.message : String(documentChoiceError), action });
      return;
    }

    const savedActions = await readFustActions();
    const actionIndex = savedActions.findIndex((item) => item.id === action.id);
    if (actionIndex >= 0) {
      savedActions[actionIndex] = action;
      await writeFustActions(savedActions);
    }
    await mirrorFustActionToDatabase(action);

    sendJson(res, 201, { action });
    return;
  }

  if (url.pathname === "/api/fust/import/preview" && req.method === "POST") {
    if (!requirePermission(res, requestUser, PERMISSIONS.FUST_MANAGE)) {
      return;
    }
    try {
      const body = await readRequestJson(req, 25 * 1024 * 1024);
      const prepared = await prepareFustImportRows(body?.file || {}, requestUser);
      sendJson(res, 200, {
        file_name: prepared.file_name,
        sheet_name: prepared.sheet_name,
        rows: prepared.rows,
        summary: {
          total_rows: prepared.rows.length,
          new_rows: prepared.rows.filter((row) => row.status === "new").length,
          update_rows: prepared.rows.filter((row) => row.status === "update").length,
          locked_rows: prepared.rows.filter((row) => row.status === "locked").length,
          missing_connect_rows: prepared.rows.filter((row) => row.status === "missing_connect").length,
        },
      });
    } catch (error) {
      sendJson(res, 400, { error: error instanceof Error ? error.message : String(error) });
    }
    return;
  }

  if (url.pathname === "/api/fust/import/apply" && req.method === "POST") {
    if (!requirePermission(res, requestUser, PERMISSIONS.FUST_MANAGE)) {
      return;
    }
    try {
      const body = await readRequestJson(req, 25 * 1024 * 1024);
      const payload = await applyFustImportRows(body?.file || {}, requestUser, body?.selected_import_keys || []);
      sendJson(res, 200, payload);
    } catch (error) {
      sendJson(res, 400, { error: error instanceof Error ? error.message : String(error) });
    }
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
    const status = normalizeSyncStatus(await readJsonFile(syncStatusPath, {}));
    if (status?.stale) {
      await writeJsonFile(syncStatusPath, status);
    }
    sendJson(res, 200, status);
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

async function startServer() {
  if (isDatabaseEnabled()) {
    try {
      await initializeDatabase();
      console.log("Postgres is connected and ready.");
    } catch (error) {
      console.error("Postgres initialization failed:", error instanceof Error ? error.message : String(error));
    }
  } else {
    console.log("Postgres is not configured yet. DATABASE_URL is not set.");
  }

  syncPendingDagFoutjesDaysToSheet().catch(() => {});
  setInterval(() => {
    syncPendingDagFoutjesDaysToSheet().catch(() => {});
  }, 15 * 60 * 1000);

  const runFustReminderCheck = async () => {
    try {
      const { settings, actions } = await loadCurrentFustActionsSnapshot();
      await maybeSendFustConfirmationReminders(actions, settings);
    } catch (error) {
      console.error("Fust confirmation reminder check failed:", error instanceof Error ? error.message : String(error));
    }
  };

  runFustReminderCheck().catch(() => {});
  setInterval(() => {
    runFustReminderCheck().catch(() => {});
  }, 15 * 60 * 1000);

  server.listen(port, host, () => {
    console.log("SnappySjaak shadow app is running.");
    console.log("Open on this PC:");
    console.log(`  http://127.0.0.1:${port}`);
    console.log("Open from another PC on the same network:");
    for (const url of lanUrls().filter((url) => !url.includes("127.0.0.1"))) {
      console.log(`  ${url}`);
    }
  });
}

startServer().catch((error) => {
  console.error(error);
  process.exit(1);
});
