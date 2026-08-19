#!/usr/bin/env node
// Run this ON THE OFFICE PC once Render is confirmed healthy again, after a
// failover where this PC was switched to "Online" and used for real work.
//
// Usage:
//   node reconcile.js                 (dry run -- reports what WOULD change, changes nothing)
//   node reconcile.js --apply         (actually merges Fust actions into Render)
//
// Required env vars:
//   RENDER_BASE_URL       e.g. https://snappysjaak-shadow.onrender.com
//   RECONCILE_USERNAME    an admin username on Render
//   RECONCILE_PASSWORD    that admin's password
//   SNAPPYSJAAK_CACHE_DIR this PC's local cache dir (defaults like the main app)
//
// What this does automatically: adds any brand-new Fust actions created on this
// PC during the outage into Render, and fills in missing document/reference
// fields on already-existing actions (never overwrites a value Render already has).
// What it deliberately does NOT do automatically: fust-settings.json (Google
// OAuth tokens, SMTP password) and shadow-users.json are only diffed and
// reported -- review the differences yourself and apply only what you approve
// via Settings > System mode > (or POST /api/backup/reconcile/settings-apply).
import { promises as fs } from "node:fs";
import path from "node:path";

const applyChanges = process.argv.includes("--apply");
const renderBaseUrl = String(process.env.RENDER_BASE_URL || "").trim().replace(/\/+$/, "");
const username = String(process.env.RECONCILE_USERNAME || "").trim();
const password = String(process.env.RECONCILE_PASSWORD || "");
const cacheDir = path.resolve(process.env.SNAPPYSJAAK_CACHE_DIR || path.join(process.cwd(), "..", "..", "..", ".cache"));

if (!renderBaseUrl || !username || !password) {
  console.error("Set RENDER_BASE_URL, RECONCILE_USERNAME, and RECONCILE_PASSWORD first.");
  process.exit(1);
}

async function readLocalJson(fileName, fallback) {
  try {
    return JSON.parse(await fs.readFile(path.join(cacheDir, fileName), "utf8"));
  } catch {
    return fallback;
  }
}

async function login() {
  const response = await fetch(`${renderBaseUrl}/api/auth/login`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ username, password }),
  });
  if (!response.ok) {
    throw new Error(`Login failed: HTTP ${response.status}`);
  }
  const setCookie = response.headers.get("set-cookie") || "";
  const cookie = setCookie.split(";")[0];
  if (!cookie) {
    throw new Error("Login succeeded but no session cookie was returned.");
  }
  return cookie;
}

async function postJson(cookie, pathName, payload) {
  const response = await fetch(`${renderBaseUrl}${pathName}`, {
    method: "POST",
    headers: { "content-type": "application/json", cookie },
    body: JSON.stringify(payload),
  });
  const body = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(`${pathName} failed: HTTP ${response.status} ${JSON.stringify(body)}`);
  }
  return body;
}

function printDiff(label, diffResult) {
  console.log(`\n--- ${label} ---`);
  if (diffResult.identical) {
    console.log("No differences.");
    return;
  }
  console.log(`${diffResult.differences.length} field(s) differ (review manually, nothing applied automatically):`);
  for (const entry of diffResult.differences) {
    console.log(`  ${entry.key}:`);
    console.log(`    Render (current): ${JSON.stringify(entry.current_value)}`);
    console.log(`    This PC (incoming): ${JSON.stringify(entry.incoming_value)}`);
  }
}

async function main() {
  console.log(`Mode: ${applyChanges ? "APPLY (will write to Render)" : "DRY RUN (nothing will be changed)"}`);
  console.log(`Cache dir: ${cacheDir}`);
  const cookie = await login();

  const localActions = (await readLocalJson("fust-actions.json", { actions: [] })).actions || [];
  const mergeResult = await postJson(cookie, "/api/backup/reconcile/merge-actions", {
    actions: localActions,
    dry_run: !applyChanges,
  });
  console.log("\n--- Fust actions ---");
  console.log(`Incoming from this PC: ${mergeResult.summary.incoming_total}`);
  console.log(`New actions ${applyChanges ? "added" : "that WOULD be added"}: ${mergeResult.summary.added}`);
  console.log(`Existing actions ${applyChanges ? "updated" : "that WOULD be updated"} (missing fields filled in): ${mergeResult.summary.updated}`);
  console.log(`Already up to date: ${mergeResult.summary.unchanged}`);

  const localSettings = await readLocalJson("fust-settings.json", {});
  const settingsDiff = await postJson(cookie, "/api/backup/reconcile/diff", { dataset: "fust_settings", incoming: localSettings });
  printDiff("fust-settings.json (OAuth tokens, SMTP, spreadsheet IDs)", settingsDiff);

  const localUsers = await readLocalJson("shadow-users.json", []);
  const usersDiff = await postJson(cookie, "/api/backup/reconcile/diff", { dataset: "shadow_users", incoming: { users: localUsers } });
  printDiff("shadow-users.json (accounts/permissions)", usersDiff);

  if (!applyChanges) {
    console.log("\nThis was a dry run. Re-run with --apply once you're happy with the above to actually merge the Fust actions.");
  }
}

main().catch((error) => {
  console.error("Reconciliation failed:", error instanceof Error ? error.message : String(error));
  process.exit(1);
});
