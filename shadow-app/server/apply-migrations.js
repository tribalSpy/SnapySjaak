#!/usr/bin/env node
// One-off script for update_standby_pc.bat: re-applies every Postgres schema
// migration against DATABASE_URL immediately after an update, so the local
// schema is guaranteed current before start_standby.bat is even run again --
// not just eventually, whenever the app next happens to restart.
import { initializeDatabase } from "./db.js";

if (!process.env.DATABASE_URL) {
  console.log("DATABASE_URL is not set in this environment -- skipping.");
  console.log("Migrations will still run automatically the next time start_standby.bat starts the app.");
  process.exit(0);
}

try {
  await initializeDatabase();
  console.log("Postgres schema is up to date.");
  process.exit(0);
} catch (error) {
  console.error("Applying migrations failed:", error instanceof Error ? error.message : String(error));
  process.exit(1);
}
