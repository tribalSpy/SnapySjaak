import { promises as fs } from "node:fs";
import path from "node:path";
import crypto from "node:crypto";

// Explicit allow-list of authoritative state, not a directory walk over cacheDir --
// this deliberately excludes rebuildable caches (run_data.json, shadow-google-images/,
// hal-locations/, fust-backups/, etc.) so nothing gets synced that doesn't need to be,
// and so a new state file added later has to be added here on purpose.
export const BACKUP_ALLOWED_FILES = [
  "shadow-users.json",
  "fust-actions.json",
  "fust-settings.json",
  "ukdocs-state.json",
  "clock-records.json",
  "expedition-stickers.json",
  "dag-foutjes.json",
  "bunches-state.json",
];

export const BACKUP_ALLOWED_DIRS = ["ukdocs-print-files"];

async function walkDir(baseDir, relDir) {
  const absDir = path.join(baseDir, relDir);
  let entries;
  try {
    entries = await fs.readdir(absDir, { withFileTypes: true });
  } catch {
    return [];
  }
  const results = [];
  for (const entry of entries) {
    const relPath = path.posix.join(relDir.split(path.sep).join("/"), entry.name);
    if (entry.isDirectory()) {
      results.push(...(await walkDir(baseDir, relPath)));
    } else if (entry.isFile()) {
      results.push(relPath);
    }
  }
  return results;
}

export async function listBackupRelativePaths(cacheDir) {
  const relativePaths = [];
  for (const file of BACKUP_ALLOWED_FILES) {
    try {
      await fs.access(path.join(cacheDir, file));
      relativePaths.push(file);
    } catch {
      // Not created yet (e.g. a module that's never been used) -- skip it.
    }
  }
  for (const dir of BACKUP_ALLOWED_DIRS) {
    relativePaths.push(...(await walkDir(cacheDir, dir)));
  }
  return relativePaths;
}

async function hashFile(absPath) {
  const data = await fs.readFile(absPath);
  return crypto.createHash("sha256").update(data).digest("hex");
}

// Hashes are cached by (path, size, mtime) so a 15-minute cycle only rehashes
// files that actually changed, instead of rehashing the whole cacheDir every time.
export async function buildBackupManifest(cacheDir, manifestCachePath) {
  const relativePaths = await listBackupRelativePaths(cacheDir);

  let hashCache = {};
  try {
    hashCache = JSON.parse(await fs.readFile(manifestCachePath, "utf8"));
  } catch {
    hashCache = {};
  }

  const files = [];
  const nextCache = {};
  for (const relPath of relativePaths) {
    const absPath = path.join(cacheDir, ...relPath.split("/"));
    const stat = await fs.stat(absPath);
    const cached = hashCache[relPath];
    const sha256 = cached && cached.size === stat.size && cached.mtime_ms === stat.mtimeMs
      ? cached.sha256
      : await hashFile(absPath);
    nextCache[relPath] = { size: stat.size, mtime_ms: stat.mtimeMs, sha256 };
    files.push({ path: relPath, size: stat.size, sha256, mtime_ms: stat.mtimeMs });
  }

  await fs.writeFile(manifestCachePath, JSON.stringify(nextCache), "utf8").catch(() => {});
  return { generated_at: new Date().toISOString(), files };
}

export function isAllowedBackupPath(relPath) {
  const normalized = String(relPath || "").trim().replace(/\\/g, "/");
  if (!normalized || normalized.includes("..") || normalized.startsWith("/")) {
    return false;
  }
  if (BACKUP_ALLOWED_FILES.includes(normalized)) {
    return true;
  }
  return BACKUP_ALLOWED_DIRS.some((dir) => normalized === dir || normalized.startsWith(`${dir}/`));
}
