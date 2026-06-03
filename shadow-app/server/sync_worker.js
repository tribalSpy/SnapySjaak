import { createWriteStream, existsSync, promises as fs } from "node:fs";
import path from "node:path";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const appRoot = path.resolve(__dirname, "..");
const repoRoot = path.resolve(appRoot, "..");
const cacheDir = path.resolve(process.env.SNAPPYSJAAK_CACHE_DIR || path.join(repoRoot, ".cache"));
const syncStatusPath = path.join(cacheDir, "index_sync_status.json");
const syncScriptPath = path.join(repoRoot, "sync_index.py");
const pollerRoot = process.env.CARGOSNAP_POLLER_ROOT || path.resolve(repoRoot, "..", "cargosnapPull");
const syncLogPath = path.join(cacheDir, "shadow-sync.log");

function parseArgs() {
  const args = process.argv.slice(2);
  const parsed = { mode: "rebuild", date: null };
  for (let index = 0; index < args.length; index += 1) {
    if (args[index] === "--mode") {
      parsed.mode = args[index + 1] || parsed.mode;
      index += 1;
    } else if (args[index] === "--date") {
      parsed.date = args[index + 1] || null;
      index += 1;
    }
  }
  if (!["rebuild", "refresh_date"].includes(parsed.mode)) {
    throw new Error(`Unsupported sync mode: ${parsed.mode}`);
  }
  if (parsed.mode === "refresh_date" && !parsed.date) {
    throw new Error("--date is required for refresh_date");
  }
  return parsed;
}

async function writeStatus(state, mode, extra = {}) {
  await fs.mkdir(cacheDir, { recursive: true });
  const payload = {
    state,
    mode,
    updated_at: new Date().toISOString(),
    ...extra,
  };
  await fs.writeFile(syncStatusPath, JSON.stringify(payload, null, 2), "utf8");
}

function localPython(projectRoot) {
  const venvPython = path.join(projectRoot, ".venv", "Scripts", "python.exe");
  if (existsSync(venvPython)) {
    return venvPython;
  }
  const unixVenvPython = path.join(projectRoot, ".venv", "bin", "python");
  if (existsSync(unixVenvPython)) {
    return unixVenvPython;
  }
  if (process.env.PYTHON) {
    return process.env.PYTHON;
  }
  return process.platform === "win32" ? "python" : "python3";
}

function appendLogHeader(stream, label, command, args, cwd) {
  stream.write(`\n[${new Date().toISOString()}] ${label}\n`);
  stream.write(`cwd: ${cwd}\n`);
  stream.write(`cmd: ${command} ${args.join(" ")}\n`);
}

function runCommand({ label, command, args, cwd }) {
  return new Promise((resolve, reject) => {
    const log = createWriteStream(syncLogPath, { flags: "a" });
    appendLogHeader(log, label, command, args, cwd);

    const child = spawn(command, args, {
      cwd,
      windowsHide: true,
      env: process.env,
    });

    child.stdout.on("data", (chunk) => log.write(chunk));
    child.stderr.on("data", (chunk) => log.write(chunk));
    child.on("error", (error) => {
      log.end();
      reject(error);
    });
    child.on("close", (code) => {
      log.write(`[${new Date().toISOString()}] ${label} exited with ${code}\n`);
      log.end();
      if (code === 0) {
        resolve();
        return;
      }
      reject(new Error(`${label} exited with code ${code}`));
    });
  });
}

async function runPollerIfAvailable(startedAt) {
  if (process.env.TRIGGER_POLLER_ON_SYNC === "0") {
    return false;
  }
  if (!existsSync(path.join(pollerRoot, "app", "poller.py"))) {
    return false;
  }

  await writeStatus("running", "poller_then_index", {
    started_at: startedAt,
    step: "poller",
    poller_root: pollerRoot,
  });
  await runCommand({
    label: "CargoSnap poller",
    command: localPython(pollerRoot),
    args: ["-m", "app.poller"],
    cwd: pollerRoot,
  });
  return true;
}

async function runIndexSync({ mode, date }, startedAt, pollerRan) {
  await writeStatus("running", pollerRan ? `poller_then_${mode}` : mode, {
    started_at: startedAt,
    step: "index",
    date,
  });

  const args = [syncScriptPath, "--mode", mode];
  if (date) {
    args.push("--date", date);
  }
  await runCommand({
    label: "Run index sync",
    command: localPython(repoRoot),
    args,
    cwd: repoRoot,
  });
}

async function main() {
  const options = parseArgs();
  const startedAt = new Date().toISOString();
  let pollerRan = false;

  try {
    pollerRan = await runPollerIfAvailable(startedAt);
    await runIndexSync(options, startedAt, pollerRan);
  } catch (error) {
    await writeStatus("failed", pollerRan ? `poller_then_${options.mode}` : options.mode, {
      started_at: startedAt,
      error: error instanceof Error ? error.message : String(error),
    });
    process.exitCode = 1;
  }
}

main();
