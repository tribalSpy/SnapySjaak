#!/usr/bin/env node
// One-time setup helper: generates every secret the office-PC standby needs
// (the X25519 backup encryption keypair + a shared agent API key) and writes
// them straight into a ready-to-fill-in start_standby.bat at the repo root,
// so nobody has to hand-copy long random strings between a console and an
// env-var editor. Run this ONCE, on the office PC.
//
// If start_standby.bat already exists, this does nothing and exits --
// re-generating would produce a new BACKUP_PRIVATE_KEY that no longer matches
// whatever BACKUP_PUBLIC_KEY was already handed to Render, silently breaking
// decryption. Delete start_standby.bat first if you deliberately want to
// rotate the keys (and update BACKUP_PUBLIC_KEY/BACKUP_AGENT_API_KEY on
// Render to match the new values this prints).
import crypto from "node:crypto";
import { existsSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { generateBackupKeyPair } from "./crypto.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(__dirname, "..", "..", "..");
const startScriptPath = path.join(repoRoot, "start_standby.bat");

if (existsSync(startScriptPath)) {
  console.log(`${startScriptPath} already exists -- secrets were already generated once.`);
  console.log("Not generating new ones (that would invalidate the public key already given to Render).");
  console.log("To deliberately rotate keys: delete start_standby.bat, re-run this script, then update");
  console.log("BACKUP_PUBLIC_KEY and BACKUP_AGENT_API_KEY on Render to the newly printed values.");
  process.exit(0);
}

const { publicKey, privateKey } = generateBackupKeyPair();
const agentApiKey = crypto.randomBytes(32).toString("hex");

console.log("Paste these two into Render's dashboard environment variables:");
console.log("");
console.log(`BACKUP_PUBLIC_KEY=${publicKey}`);
console.log(`BACKUP_AGENT_API_KEY=${agentApiKey}`);
console.log("");
console.log("BACKUP_PRIVATE_KEY stays on this PC only -- it has already been written into");
console.log(`${startScriptPath}. Also save a copy of it somewhere safe (e.g. a`);
console.log("password manager) -- losing it means losing every backup ever taken:");
console.log("");
console.log(privateKey);

const template = `@echo off
title SnappySjaak Standby
cd /d "%~dp0"

REM --- Fill these two in before first use ---
set RENDER_BACKUP_BASE_URL=https://REPLACE-WITH-YOUR-RENDER-URL.onrender.com
set DATABASE_URL=REPLACE-WITH-YOUR-LOCAL-POSTGRES-CONNECTION-STRING

REM --- Generated once by generate-keypair.js -- keep this file private, never commit it ---
set BACKUP_AGENT_API_KEY=${agentApiKey}
set BACKUP_PRIVATE_KEY=${privateKey}

REM --- Adjust these if you want different locations ---
set SNAPPYSJAAK_CACHE_DIR=%~dp0standby-cache
set GOOGLE_APPLICATION_CREDENTIALS=%~dp0credentials\\service_account.json
set PORT=4174

cd /d "%~dp0shadow-app"
call npm start
pause
`;

writeFileSync(startScriptPath, template, "utf8");
console.log("");
console.log(`Wrote ${startScriptPath} with BACKUP_AGENT_API_KEY and BACKUP_PRIVATE_KEY already filled in.`);
console.log("Before running it: fill in RENDER_BACKUP_BASE_URL and DATABASE_URL, and make sure");
console.log("credentials/service_account.json exists (copy it from your local dev setup).");
