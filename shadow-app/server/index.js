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
const syncScriptPath = path.join(repoRoot, "sync_index.py");
const driveBridgePath = path.join(appRoot, "server", "drive_bridge.py");
const syncWorkerPath = path.join(appRoot, "server", "sync_worker.js");
const googleImageCacheDir = path.join(cacheDir, "shadow-google-images");
const staticRoot = existsSync(path.join(appRoot, "dist"))
  ? path.join(appRoot, "dist")
  : path.join(appRoot, "public");
const autoSyncOnVisit = process.env.AUTO_SYNC_ON_VISIT !== "0";
const autoSyncThrottleMs = Number(process.env.AUTO_SYNC_THROTTLE_MINUTES || 5) * 60 * 1000;
const autoSyncStartedAt = new Map();
const sessions = new Map();
const sessionCookieName = "snappysjaak_shadow_session";

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

async function readRequestJson(req) {
  const chunks = [];
  for await (const chunk of req) {
    chunks.push(chunk);
    if (Buffer.concat(chunks).length > 1024 * 1024) {
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

async function readUsers() {
  const payload = await readJsonFile(usersPath, { users: [] });
  return Array.isArray(payload?.users) ? payload.users : [];
}

async function writeUsers(users) {
  await writeJsonFile(usersPath, { users });
}

function publicUser(user) {
  return {
    username: user.username,
    role: user.role,
    created_at: user.created_at,
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

      reject(new Error(Buffer.concat(stderr).toString("utf8") || `Python bridge exited with ${code}`));
    });

    if (input) {
      child.stdin.write(input);
    }
    child.stdin.end();
  });
}

async function hydrateGoogleRuns(runs) {
  const googleRuns = runs.filter((run) => String(run?.metadata?.source || "google_drive") === "google_drive");
  if (!googleRuns.length) {
    return new Map();
  }

  const output = await runPythonBridge(["details"], JSON.stringify(googleRuns));
  const hydratedRuns = JSON.parse(output.toString("utf8"));
  return new Map(hydratedRuns.map((run) => [run.folder_id, run]));
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

async function readGoogleImage(fileId, accountName = "default") {
  const cachePath = googleImageCachePath(fileId, accountName);
  if (existsSync(cachePath)) {
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
    if (!imagePath) {
      sendText(res, 400, "Invalid image id");
      return;
    }

    if (!path.isAbsolute(imagePath)) {
      try {
        const imageBytes = await readGoogleImage(imagePath, accountName);
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
