#!/usr/bin/env node
// One-time setup helper: generates the X25519 keypair used to encrypt/decrypt
// backup payloads. Run this ONCE, on the office PC, and:
//   - put BACKUP_PUBLIC_KEY on Render (render.yaml / dashboard env var) --
//     it only lets Render seal data, never read it back.
//   - put BACKUP_PRIVATE_KEY only on the office PC, never on Render, never in git.
//     Losing it means losing the ability to decrypt every backup ever taken,
//     so also store a copy somewhere safe (e.g. a password manager).
import { generateBackupKeyPair } from "./crypto.js";

const { publicKey, privateKey } = generateBackupKeyPair();

console.log("BACKUP_PUBLIC_KEY (set this on Render):");
console.log(publicKey);
console.log("");
console.log("BACKUP_PRIVATE_KEY (office PC only -- never on Render, never in git):");
console.log(privateKey);
