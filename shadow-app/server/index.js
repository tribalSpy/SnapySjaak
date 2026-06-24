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
const ukdocsStatePath = path.join(cacheDir, "ukdocs-state.json");
const fustBackupDir = path.join(cacheDir, "fust-backups");
const syncScriptPath = path.join(repoRoot, "sync_index.py");
const driveBridgePath = path.join(appRoot, "server", "drive_bridge.py");
const syncWorkerPath = path.join(appRoot, "server", "sync_worker.js");
const halLocationsWorkerPath = path.join(appRoot, "server", "hal_locations_worker.py");
const expeditionStickerWorkerPath = path.join(appRoot, "server", "expedition_sticker_worker.py");
const ukdocsWorkerPath = path.join(appRoot, "server", "ukdocs_worker.py");
const googleImageCacheDir = path.join(cacheDir, "shadow-google-images");
const googleRunDetailsCacheDir = path.join(cacheDir, "shadow-google-run-details");
const halLocationsCacheDir = path.join(cacheDir, "hal-locations");
const expeditionStickerStatePath = path.join(cacheDir, "expedition-stickers.json");
const expeditionStickerFilesDir = path.join(cacheDir, "expedition-stickers");
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
const halLocationSessions = new Map();
const sessionCookieName = "snappysjaak_shadow_session";
const halLocationSessionTtlMs = 30 * 60 * 1000;
const allPermissions = [
  "photos:view",
  "fust:view",
  "fust:in",
  "fust:out",
  "fust:overview",
  "cmr:view",
  "hal_locations:view",
  "expedition_stickers:view",
  "cmr:manage",
  "clock:view",
  "clock:manage",
  "users:manage",
  "settings:manage",
  "ukdocs:view",
];
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
  hal_locations_spreadsheet_id: "",
  hal_locations_sheet_name: "ERP_PASTE",
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
};

const defaultExpeditionStickerState = {
  planning_file: null,
  split_file: null,
};

const cmrPrintDataDirCandidates = [
  path.join(repoRoot, "cmrprint", "CMRPrint", "bin", "Release", "net9.0-windows", "win-x64", "publish", "Data"),
  path.join(repoRoot, "cmrprint", "CMRPrint", "bin", "Release", "net9.0-windows", "win-x64", "Data"),
  path.join(repoRoot, "cmrprint", "CMRPrint", "Data"),
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
  const { dataDir, candidates } = resolveCmrPrintDataDir();
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
  return path.join(repoRoot, "cmrprint", "CMRPrint", "Data");
}

async function ensureCmrPrintDataDir() {
  const dataDir = resolveCmrPrintDataDir().dataDir || cmrPrintPrimaryDataDir();
  const templatesDir = path.join(dataDir, "Templates");
  await fs.mkdir(templatesDir, { recursive: true });
  return { dataDir, templatesDir, appDataPath: path.join(dataDir, "app-data.xml") };
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
    hal_locations_spreadsheet_id: String(settings?.hal_locations_spreadsheet_id || settings?.spreadsheet_id || "").trim(),
    hal_locations_sheet_name: String(settings?.hal_locations_sheet_name || defaultFustSettings.hal_locations_sheet_name).trim() || defaultFustSettings.hal_locations_sheet_name,
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

function normalizeUkdocsText(value) {
  return String(value || "").trim();
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
    default_invoice_language_text: String(customer?.default_invoice_language_text || "").trim(),
    default_document_references: String(customer?.default_document_references || "").trim(),
    show_invoice_vat_number: customer?.show_invoice_vat_number !== false,
    show_invoice_eori_number: customer?.show_invoice_eori_number !== false,
    show_invoice_importer_number: customer?.show_invoice_importer_number !== false,
    export_defaults: normalizeUkdocsExportDefaults(customer?.export_defaults || {}),
  };
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
    created_at: normalizeUkdocsText(report?.created_at),
    final_status: normalizeUkdocsText(report?.final_status),
    warnings: Array.isArray(report?.warnings) ? report.warnings.map((item) => String(item || "").trim()).filter(Boolean) : [],
    summary: String(report?.summary || "").trim(),
  };
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
    deleted: action?.deleted === true,
    deleted_at: String(action?.deleted_at || ""),
    deleted_by: String(action?.deleted_by || ""),
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

function mergeUkdocsStatePatch(currentState, patch) {
  return normalizeUkdocsState({
    ...currentState,
    ...patch,
    company_settings: patch.company_settings ? { ...currentState.company_settings, ...patch.company_settings } : currentState.company_settings,
    export_defaults: patch.export_defaults ? { ...currentState.export_defaults, ...patch.export_defaults } : currentState.export_defaults,
  });
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

async function syncFustActionToSheets(action, settings, options = {}) {
  if (!settings.spreadsheet_id) {
    return { ok: false, target_sheets: [], error: "Spreadsheet ID is not configured" };
  }

  const targetSheet = action.type === "OUT" ? settings.out_sheet_name : settings.in_sheet_name;
  if (!targetSheet) {
    return { ok: false, target_sheets: [], error: "Target sheet is not configured" };
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
    if (!requirePermission(res, requestUser, PERMISSIONS.UKDOCS_VIEW)) {
      return;
    }

    if (req.method === "GET") {
      const state = await readUkdocsState();
      sendJson(res, 200, { state });
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
    const output = await runUkdocsWorker(["generate"], JSON.stringify(body));
    sendJson(res, 200, JSON.parse(output.toString("utf8")));
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
    const nextShipments = state.shipments.filter((item) => item.id !== shipmentId);
    if (nextShipments.length === state.shipments.length) {
      sendJson(res, 404, { error: "Shipment not found" });
      return;
    }
    state.shipments = nextShipments;
    await writeUkdocsState(state);
    sendJson(res, 200, { ok: true, shipments: normalizeUkdocsState(state).shipments });
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

    const settings = await readFustSettings();
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

    const deletedAction = normalizeFustAction({
      ...action,
      deleted: true,
      deleted_at: new Date().toISOString(),
      deleted_by: requestUser.username,
      sheet_sync: {
        ...(action.sheet_sync || {}),
        ok: false,
        error: "Deleted locally",
      },
      email_sync: {
        ...(action.email_sync || {}),
        ok: false,
        error: "Deleted locally",
      },
    });
    if (actionIndex >= 0) {
      actions[actionIndex] = deletedAction;
    } else {
      actions.push(deletedAction);
    }
    await writeFustActions(actions);
    sendJson(res, 200, { ok: true, deleted_action_id: actionId });
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
