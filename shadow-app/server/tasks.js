// Tasks + chat between users (the floating widget).
//
// Tasks are given by hand to a person or a team, or made by the app itself
// (the PD step plan, important invoice differences, follow-up lines moved to
// a department, claims solved). Chat: one-on-one (dm:a|b), per team
// (team:<key>) and per task (task:<id>). Everything lives in its own tables
// (app_tasks, app_task_events, app_chat_messages, app_chat_reads,
// app_task_settings); the users file only gets each user's teams and
// widget settings. Nothing here changes how the rest of the app works: every
// hook from elsewhere is best effort and never throws back.
import crypto from "node:crypto";
import { dbQuery, isDatabaseEnabled } from "./db.js";

export const APP_TASK_DEFAULT_TEAMS = [
  { key: "inkoop", name: "Inkoop" },
  { key: "invoice_control", name: "Invoice control" },
  { key: "finance", name: "Finance" },
  { key: "expeditie", name: "Expeditie" },
  { key: "pd_phyto", name: "PD / Phyto" },
];
const TASK_STATUSES = ["open", "in_progress", "done", "cancelled"];
const TASK_PRIORITIES = ["normal", "high", "urgent"];
const TASK_REPEATS = ["none", "daily", "workdays", "weekly"];
const OPEN_STATUSES = ["open", "in_progress"];

export function normalizeTeamKey(value) {
  return String(value || "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "").slice(0, 40);
}

export function normalizeUserTeams(value) {
  return [...new Set((Array.isArray(value) ? value : []).map(normalizeTeamKey).filter(Boolean))].slice(0, 20);
}

export function normalizeUserWidgetSettings(value) {
  const source = value && typeof value === "object" && !Array.isArray(value) ? value : {};
  const pos = source.widget_pos && typeof source.widget_pos === "object" ? source.widget_pos : null;
  const x = Number(pos?.x);
  const y = Number(pos?.y);
  return {
    use_widget: source.use_widget === true,
    widget_pos: pos && Number.isFinite(x) && Number.isFinite(y)
      ? { x: Math.max(0, Math.min(1, x)), y: Math.max(0, Math.min(1, y)) }
      : null,
  };
}

function isoDay(value) {
  const text = String(value || "").trim().slice(0, 10);
  return /^\d{4}-\d{2}-\d{2}$/.test(text) ? text : "";
}

function localToday() {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;
}

function addDays(day, count) {
  const date = new Date(`${day}T12:00:00Z`);
  date.setUTCDate(date.getUTCDate() + count);
  return date.toISOString().slice(0, 10);
}

function nextRepeatDay(repeat, fromDay) {
  const base = isoDay(fromDay) || localToday();
  if (repeat === "daily") return addDays(base, 1);
  if (repeat === "weekly") return addDays(base, 7);
  if (repeat === "workdays") {
    let next = addDays(base, 1);
    while ([0, 6].includes(new Date(`${next}T12:00:00Z`).getUTCDay())) next = addDays(next, 1);
    return next;
  }
  return "";
}

// "user:name" / "team:key" / "" -> { assignee_user, assignee_team }
function parseAssignee(value) {
  const text = String(value || "").trim();
  if (text.startsWith("team:")) return { assignee_user: null, assignee_team: normalizeTeamKey(text.slice(5)) || null };
  if (text.startsWith("user:")) return { assignee_user: text.slice(5).trim() || null, assignee_team: null };
  return { assignee_user: null, assignee_team: null };
}

function dmChannel(left, right) {
  return `dm:${[String(left), String(right)].sort().join("|")}`;
}

function taskRow(row) {
  if (!row) return null;
  return {
    id: String(row.id),
    title: row.title,
    body: row.body || "",
    assignee_user: row.assignee_user || "",
    assignee_team: row.assignee_team || "",
    created_by: row.created_by || "",
    priority: row.priority,
    due_date: row.due_date ? (row.due_date instanceof Date ? `${row.due_date.getFullYear()}-${String(row.due_date.getMonth() + 1).padStart(2, "0")}-${String(row.due_date.getDate()).padStart(2, "0")}` : String(row.due_date).slice(0, 10)) : "",
    status: row.status,
    done_by: row.done_by || "",
    done_at: row.done_at || null,
    source: row.source,
    link: row.link || {},
    repeat: row.repeat || "none",
    created_at: row.created_at,
    updated_at: row.updated_at,
    messages: row.messages === undefined ? undefined : Number(row.messages) || 0,
  };
}

// ---------------------------------------------------------------------------
// Settings (teams list, automatic task rules) -- one jsonb row.

function normalizeTaskConfig(value) {
  const source = value && typeof value === "object" ? value : {};
  const teams = (Array.isArray(source.teams) && source.teams.length ? source.teams : APP_TASK_DEFAULT_TEAMS)
    .map((team) => ({ key: normalizeTeamKey(team?.key || team?.name), name: String(team?.name || team?.key || "").trim().slice(0, 60) }))
    .filter((team, index, list) => team.key && team.name && list.findIndex((other) => other.key === team.key) === index)
    .slice(0, 40);
  const assign = {};
  for (const [key, target] of Object.entries(source.step_plan?.assign || {})) {
    const cleanKey = String(key || "").trim().slice(0, 60);
    const cleanTarget = String(target || "").trim().slice(0, 80);
    if (cleanKey && /^(user|team):./.test(cleanTarget)) assign[cleanKey] = cleanTarget;
  }
  // One step plan task per day, for this team / person (the older per-step
  // choice: its first one).
  const assignTo = String(source.step_plan?.assign_to ?? Object.values(assign)[0] ?? "").trim().slice(0, 80);
  const minAmount = Number(source.invoice_diff?.min_amount);
  const departmentTeam = { control: "invoice_control", inkoop: "inkoop", finance: "finance", ...(source.followup?.department_team || {}) };
  return {
    teams,
    step_plan: { enabled: source.step_plan?.enabled !== false, assign_to: /^(user|team):./.test(assignTo) ? assignTo : "" },
    invoice_diff: {
      enabled: source.invoice_diff?.enabled !== false,
      min_amount: Number.isFinite(minAmount) && minAmount > 0 ? Math.round(minAmount * 100) / 100 : 50,
      team: normalizeTeamKey(source.invoice_diff?.team) || "invoice_control",
      days_back: Math.max(1, Math.min(31, Number(source.invoice_diff?.days_back) || 7)),
    },
    followup: {
      enabled: source.followup?.enabled !== false,
      department_team: Object.fromEntries(Object.entries(departmentTeam).map(([key, team]) => [key, normalizeTeamKey(team)])),
    },
    claims: { enabled: source.claims?.enabled !== false },
    // The same factuur number on more than one sending: a task per number
    // for this team to check and fix (Edit invoice numbers).
    invoice_duplicates: {
      enabled: source.invoice_duplicates?.enabled !== false,
      team: normalizeTeamKey(source.invoice_duplicates?.team) || "pd_phyto",
    },
    // A daily task with a "Run on this PC" button that starts a Uniware
    // recorder action (the facturen for Finance Audit).
    facturen: {
      enabled: source.facturen?.enabled === true,
      assign_to: /^(user|team):./.test(String(source.facturen?.assign_to || "")) ? String(source.facturen.assign_to).trim().slice(0, 80) : "",
      action: String(source.facturen?.action || "facturen-finance-audit").trim().slice(0, 80) || "facturen-finance-audit",
    },
    updated_at: source.updated_at || "",
    updated_by: source.updated_by || "",
  };
}

let configCache = null;
export async function readTaskConfig() {
  if (!isDatabaseEnabled()) return normalizeTaskConfig(null);
  if (configCache && Date.now() - configCache.at < 30000) return configCache.value;
  const result = await dbQuery("SELECT value FROM app_task_settings WHERE key = 'config'");
  const value = normalizeTaskConfig(result.rows[0]?.value);
  configCache = { at: Date.now(), value };
  return value;
}

async function writeTaskConfig(value, by) {
  const normalized = normalizeTaskConfig({ ...value, updated_at: new Date().toISOString(), updated_by: by });
  await dbQuery(
    `INSERT INTO app_task_settings (key, value, updated_by, updated_at) VALUES ('config', $1, $2, now())
     ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value, updated_by = EXCLUDED.updated_by, updated_at = now()`,
    [JSON.stringify(normalized), by || ""],
  );
  configCache = null;
  return normalized;
}

// ---------------------------------------------------------------------------
// Tasks

async function logTaskEvent(taskId, by, action, from = "", to = "") {
  await dbQuery(
    "INSERT INTO app_task_events (task_id, by_user, action, from_value, to_value) VALUES ($1, $2, $3, $4, $5)",
    [taskId, by || "", action, String(from ?? "").slice(0, 2000), String(to ?? "").slice(0, 2000)],
  );
}

async function insertTask(fields, by) {
  const result = await dbQuery(
    `INSERT INTO app_tasks (title, body, assignee_user, assignee_team, created_by, priority, due_date, source, source_key, link, repeat)
     VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11)
     ON CONFLICT (source_key) DO NOTHING
     RETURNING *`,
    [
      String(fields.title || "").trim().slice(0, 300) || "Task",
      String(fields.body || "").slice(0, 8000),
      fields.assignee_user || null,
      fields.assignee_team || null,
      fields.created_by || by || "",
      TASK_PRIORITIES.includes(fields.priority) ? fields.priority : "normal",
      isoDay(fields.due_date) || null,
      fields.source || "manual",
      fields.source_key || null,
      JSON.stringify(fields.link && typeof fields.link === "object" ? fields.link : {}),
      TASK_REPEATS.includes(fields.repeat) ? fields.repeat : "none",
    ],
  );
  const row = result.rows[0] || null;
  if (row) await logTaskEvent(row.id, by || fields.created_by || "system", "created", "", fields.assignee_user ? `user:${fields.assignee_user}` : fields.assignee_team ? `team:${fields.assignee_team}` : "");
  return row;
}

// A task the app makes by itself: made once per source_key (a day compared
// again, a step plan checked again, never doubles it). Never throws.
export async function createSystemTask(fields) {
  if (!isDatabaseEnabled() || !fields?.source_key) return null;
  try {
    return await insertTask({ ...fields, created_by: fields.created_by || "system" }, fields.created_by || "system");
  } catch (error) {
    console.error("Auto task not made:", fields.source_key, error?.message || error);
    return null;
  }
}

async function closeTasks(rows, by, note) {
  for (const row of rows) {
    await dbQuery("UPDATE app_tasks SET status = 'done', done_by = $2, done_at = now(), updated_at = now() WHERE id = $1 AND status IN ('open', 'in_progress')", [row.id, by]);
    await logTaskEvent(row.id, by, "status", row.status, "done");
    if (note) await logTaskEvent(row.id, by, "note", "", note);
  }
}

function canChangeTask(task, user) {
  if (!task || !user) return false;
  if (user.role === "admin") return true;
  const teams = user.teams || [];
  return task.created_by === user.username
    || task.assignee_user === user.username
    || (task.assignee_team && teams.includes(task.assignee_team));
}

const PRIORITY_ORDER = "CASE priority WHEN 'urgent' THEN 0 WHEN 'high' THEN 1 ELSE 2 END";

async function listTasks(user, scope) {
  const me = user.username;
  const teams = user.teams || [];
  const visible = "(assignee_user = $1 OR assignee_team = ANY($2::text[]) OR created_by = $1)";
  const messages = "(SELECT count(*) FROM app_chat_messages m WHERE m.channel = 'task:' || app_tasks.id AND NOT m.deleted) AS messages";
  let where;
  let limit = 200;
  if (scope === "team") where = `status IN ('open', 'in_progress') AND assignee_team = ANY($2::text[])`;
  else if (scope === "created") where = `status IN ('open', 'in_progress') AND created_by = $1 AND coalesce(assignee_user, '') <> $1`;
  else if (scope === "done") { where = `status IN ('done', 'cancelled') AND ${visible} AND updated_at > now() - interval '30 days'`; limit = 100; }
  // All tasks = the user's own and their teams' together (not the ones they
  // gave to others).
  else if (scope === "all") { where = "status IN ('open', 'in_progress') AND (assignee_user = $1 OR assignee_team = ANY($2::text[]))"; limit = 300; }
  // The user's tab: their own tasks and the ones they asked of others.
  else if (scope === "user") where = `status IN ('open', 'in_progress') AND (assignee_user = $1 OR (created_by = $1 AND coalesce(assignee_user, '') <> $1))`;
  else where = `status IN ('open', 'in_progress') AND assignee_user = $1`;
  // Both parameters typed in every scope (one scope uses only one of them).
  const result = await dbQuery(
    `SELECT *, ${messages} FROM app_tasks WHERE $1::text IS NOT NULL AND $2::text[] IS NOT NULL AND ${where}
     ORDER BY ${PRIORITY_ORDER}, due_date NULLS LAST, id DESC LIMIT ${limit}`,
    [me, teams],
  );
  const counts = await dbQuery(
    `SELECT
       count(*) FILTER (WHERE assignee_user = $1) AS mine,
       count(*) FILTER (WHERE assignee_team = ANY($2::text[])) AS team,
       count(*) FILTER (WHERE created_by = $1 AND coalesce(assignee_user, '') <> $1) AS created,
       count(*) FILTER (WHERE assignee_user = $1 OR assignee_team = ANY($2::text[])) AS all_open
     FROM app_tasks WHERE status IN ('open', 'in_progress')`,
    [me, teams],
  );
  const row = counts.rows[0] || {};
  return {
    tasks: result.rows.map(taskRow),
    counts: {
      mine: Number(row.mine) || 0,
      team: Number(row.team) || 0,
      created: Number(row.created) || 0,
      user: (Number(row.mine) || 0) + (Number(row.created) || 0),
      all: Number(row.all_open) || 0,
    },
  };
}

async function getTask(id) {
  if (!/^\d+$/.test(String(id || ""))) return null;
  const result = await dbQuery("SELECT * FROM app_tasks WHERE id = $1", [id]);
  return result.rows[0] || null;
}

function canSeeTask(task, user) {
  return canChangeTask(task, user);
}

// ---------------------------------------------------------------------------
// Chat

function channelKind(channel) {
  const text = String(channel || "");
  if (text.startsWith("dm:")) return "dm";
  if (text.startsWith("team:")) return "team";
  if (text.startsWith("task:")) return "task";
  return "";
}

async function canUseChannel(channel, user, users) {
  const kind = channelKind(channel);
  if (kind === "dm") {
    const pair = channel.slice(3).split("|");
    if (pair.length !== 2 || !pair.includes(user.username)) return false;
    const other = pair[0] === user.username ? pair[1] : pair[0];
    return users.some((item) => item.username === other) || other === user.username;
  }
  if (kind === "team") {
    const key = channel.slice(5);
    return user.role === "admin" || (user.teams || []).includes(key);
  }
  if (kind === "task") {
    const task = await getTask(channel.slice(5));
    return Boolean(task) && canSeeTask(task, user);
  }
  return false;
}

// The channels a user has: their DMs and task threads with messages, and
// every team of theirs -- each with its unread count and last message.
async function listChannels(user, config) {
  const me = user.username;
  const teams = user.teams || [];
  const teamChannels = teams.map((key) => `team:${key}`);
  const result = await dbQuery(
    `WITH mine AS (
       SELECT m.channel, max(m.id) AS last_id,
              count(*) FILTER (WHERE m.id > coalesce(r.last_read_id, 0) AND m.from_user <> $1 AND NOT m.deleted) AS unread
       FROM app_chat_messages m
       LEFT JOIN app_chat_reads r ON r.channel = m.channel AND r.user_name = $1
       WHERE (m.channel LIKE 'dm:%' AND $1 = ANY(string_to_array(substr(m.channel, 4), '|')))
          OR m.channel = ANY($2::text[])
          OR (m.channel LIKE 'task:%' AND EXISTS (
                SELECT 1 FROM app_tasks t WHERE 'task:' || t.id = m.channel
                  AND (t.assignee_user = $1 OR t.assignee_team = ANY($3::text[]) OR t.created_by = $1)))
       GROUP BY m.channel
     )
     SELECT mine.channel, mine.last_id, mine.unread, last.from_user, last.body, last.at, last.deleted
     FROM mine JOIN app_chat_messages last ON last.id = mine.last_id`,
    [me, teamChannels, teams],
  );
  const byChannel = new Map(result.rows.map((row) => [row.channel, row]));
  const taskIds = result.rows.filter((row) => row.channel.startsWith("task:")).map((row) => row.channel.slice(5));
  const taskTitles = new Map();
  if (taskIds.length) {
    const tasks = await dbQuery("SELECT id, title, status FROM app_tasks WHERE id = ANY($1::bigint[])", [taskIds]);
    for (const row of tasks.rows) taskTitles.set(String(row.id), row);
  }
  const teamName = (key) => config.teams.find((team) => team.key === key)?.name || key;
  const channels = [];
  const seen = new Set();
  const push = (channel) => {
    if (seen.has(channel)) return;
    seen.add(channel);
    const row = byChannel.get(channel);
    const kind = channelKind(channel);
    let title = channel;
    let with_user = "";
    if (kind === "dm") {
      const pair = channel.slice(3).split("|");
      with_user = pair[0] === me ? pair[1] : pair[0];
      title = with_user;
    } else if (kind === "team") title = teamName(channel.slice(5));
    else if (kind === "task") title = taskTitles.get(channel.slice(5))?.title || `Task ${channel.slice(5)}`;
    channels.push({
      channel,
      kind,
      title,
      with_user,
      task_id: kind === "task" ? channel.slice(5) : "",
      task_status: kind === "task" ? taskTitles.get(channel.slice(5))?.status || "" : "",
      unread: Number(row?.unread) || 0,
      last_id: row ? String(row.last_id) : "0",
      last: row ? { from: row.from_user, body: row.deleted ? "(deleted)" : String(row.body || "").slice(0, 140), at: row.at } : null,
    });
  };
  for (const row of result.rows) push(row.channel);
  for (const channel of teamChannels) push(channel);
  channels.sort((left, right) => Number(right.last_id) - Number(left.last_id) || left.title.localeCompare(right.title));
  return channels;
}

function messageRow(row) {
  return {
    id: String(row.id),
    channel: row.channel,
    from: row.from_user,
    body: row.deleted ? "" : row.body,
    deleted: row.deleted === true,
    at: row.at,
    edited_at: row.edited_at || null,
  };
}

// ---------------------------------------------------------------------------
// Automatic tasks

// Follow-up lines moved to a department: one task per grower for that
// department's team.
export async function autoTasksForFollowupMove({ department, departmentLabel, entries, note, by }) {
  try {
    if (!isDatabaseEnabled() || !department || !entries?.length) return;
    const config = await readTaskConfig();
    if (!config.followup.enabled) return;
    const team = config.followup.department_team[department];
    if (!team) return;
    const byGrower = new Map();
    for (const entry of entries) {
      const grower = String(entry.grower || "").trim() || "?";
      byGrower.set(grower, [...(byGrower.get(grower) || []), entry]);
    }
    for (const [grower, list] of byGrower) {
      const ids = list.map((entry) => String(entry.id)).sort();
      const hash = crypto.createHash("sha1").update(ids.join(",")).digest("hex").slice(0, 12);
      const money = list.reduce((sum, entry) => sum + (Number(entry.money_diff) || 0), 0);
      const days = [...new Set(list.map((entry) => entry.day).filter(Boolean))].sort();
      await createSystemTask({
        source: "followup",
        source_key: `followup:${department}:${hash}`,
        title: `Follow-up for ${departmentLabel || department}: ${grower} -- ${list.length} line(s)${money ? `, € ${money.toFixed(2)}` : ""}`,
        body: [note ? `Reason: ${note}` : "", days.length ? `Day(s): ${days.join(", ")}` : "", `Sent by ${by || "?"}`].filter(Boolean).join("\n"),
        assignee_team: team,
        created_by: by || "system",
        priority: Math.abs(money) >= 500 ? "high" : "normal",
        link: { page: "inkoop", tab: "fu", grower, item_ids: ids },
      });
    }
  } catch (error) {
    console.error("Follow-up task not made:", error?.message || error);
  }
}

// Follow-up lines resolved: their tasks close once all their lines are.
export async function autoCloseFollowupTasks(resolvedIds, isResolved, by) {
  try {
    if (!isDatabaseEnabled() || !resolvedIds?.length) return;
    const result = await dbQuery(
      "SELECT * FROM app_tasks WHERE source = 'followup' AND status IN ('open', 'in_progress') AND (link -> 'item_ids') ?| $1::text[]",
      [resolvedIds.map(String)],
    );
    const done = [];
    for (const row of result.rows) {
      const ids = Array.isArray(row.link?.item_ids) ? row.link.item_ids : [];
      if (ids.length && (await isResolved(ids))) done.push(row);
    }
    await closeTasks(done, by || "system", "All its follow-up lines are resolved.");
  } catch (error) {
    console.error("Follow-up tasks not closed:", error?.message || error);
  }
}

// A claim credited / closed / written off: a task for whoever made it.
export async function autoTaskForClaimStep({ claim, step, by, note }) {
  try {
    if (!isDatabaseEnabled() || !claim?.id) return;
    const config = await readTaskConfig();
    if (!config.claims.enabled) return;
    const owner = String(claim.created_by || "").trim();
    if (!owner || owner === by) return;
    const label = step === "credit_received" ? "credit received" : step === "written_off" ? "written off" : "closed";
    const credited = claim.amount_credited !== null && claim.amount_credited !== undefined ? ` · credited € ${Number(claim.amount_credited).toFixed(2)}` : "";
    await createSystemTask({
      source: "claim",
      source_key: `claim:${claim.id}:${step}`,
      title: `Claim C-${claim.id} ${claim.grower || ""}: ${label}`,
      body: [`Claimed € ${Number(claim.amount_claimed || 0).toFixed(2)}${credited}`, claim.credit_note_number ? `Credit note ${claim.credit_note_number}` : "", note ? `Note: ${note}` : "", `By ${by || "?"}`].filter(Boolean).join("\n"),
      assignee_user: owner,
      created_by: by || "system",
      link: { page: "inkoop", tab: "fu", grower: claim.grower || "", claim_id: String(claim.id) },
    });
  } catch (error) {
    console.error("Claim task not made:", error?.message || error);
  }
}

// Important invoice differences of the ready, compared days: one task per
// grower and day whose open lines differ by at least the amount set.
export async function syncInvoiceDiffTasks({ dates, items }) {
  if (!isDatabaseEnabled()) return { made: 0, closed: 0 };
  const config = await readTaskConfig();
  if (!config.invoice_diff.enabled || !dates?.length) return { made: 0, closed: 0 };
  const groups = new Map();
  for (const item of items || []) {
    const day = isoDay(item.day || item.date);
    if (!day) continue;
    const grower = String(item.grower || "").trim() || "?";
    const key = `${day}|${grower}`;
    const group = groups.get(key) || { day, grower, lines: 0, money: 0 };
    group.lines += 1;
    group.money += Number(item.money_diff) || 0;
    groups.set(key, group);
  }
  const keep = new Set();
  let made = 0;
  for (const group of groups.values()) {
    if (Math.abs(group.money) < config.invoice_diff.min_amount) continue;
    const sourceKey = `invoice_diff:${group.day}:${group.grower}`;
    keep.add(sourceKey);
    const row = await createSystemTask({
      source: "invoice_diff",
      source_key: sourceKey,
      title: `Factuur check ${group.day}: ${group.grower} -- € ${group.money.toFixed(2)} (${group.lines} line(s))`,
      body: `Open invoice differences of ${group.grower} on ${group.day}: € ${group.money.toFixed(2)} over ${group.lines} line(s) (more than 0 = paid more than Connect counts).`,
      assignee_team: config.invoice_diff.team,
      priority: Math.abs(group.money) >= config.invoice_diff.min_amount * 5 ? "high" : "normal",
      due_date: localToday(),
      link: { page: "inkoop", tab: "fu", grower: group.grower, day: group.day },
    });
    if (row) made += 1;
  }
  // Below the amount now (lines resolved / matched since): done by itself.
  const open = await dbQuery(
    "SELECT * FROM app_tasks WHERE source = 'invoice_diff' AND status IN ('open', 'in_progress') AND split_part(source_key, ':', 2) = ANY($1::text[])",
    [dates],
  );
  const stale = open.rows.filter((row) => !keep.has(row.source_key));
  await closeTasks(stale, "system", "The open differences of this grower and day are below the amount now.");
  return { made, closed: stale.length };
}

// The PD step plan: ONE task per day for whoever the step plan is given to.
// It carries the day's progress (link.progress): the day step it is at and,
// per shipment, the step and a simple status (as the Pipeline Monitor, with
// what is stuck) -- kept up to date after every mark / skip and every 5 min.
// Done by itself when the day and every shipment are complete; opened again
// when a shipment is added later.
function stepPlanProgress(date, plan, monitor) {
  const day = plan.day || { steps: [], current_step: 0, complete: true };
  const dayStep = day.complete ? null : day.steps[day.current_step];
  const monitorById = new Map((monitor?.rows || []).map((row) => [row.collection_id, row]));
  const rows = (plan.rows || []).map((row) => {
    const step = row.complete ? null : row.steps[row.current_step];
    const watch = monitorById.get(row.collection_id);
    const stuck = (watch?.attention_reasons || []).find((reason) => /stuck/i.test(reason)) || "";
    const status = row.complete
      ? "Complete"
      : step?.key === "day"
        ? `Waiting for the day steps (${dayStep ? `${dayStep.number}. ${dayStep.label}` : "-"})`
        : step?.status || (step?.missing || [])[0] || "";
    // What the task follows: the Pipeline Monitor's four checks, done by the
    // app's own actions -- complete, CSI done, papers sent, no problem. (Eric
    // Docs has no CSI; papers turned off for the customer count as done.)
    const checks = {
      complete: watch ? watch.complete === true : row.complete === true,
      csi: row.route === "ericdocs" ? true : watch?.csi_status === "done",
      papers: watch ? (watch.papers_sent === true || watch.papers_status === "off") : false,
      ok: watch ? watch.needs_attention !== true : false,
    };
    const checksDone = Object.values(checks).filter(Boolean).length;
    return {
      id: row.collection_id,
      checks,
      pct: Math.round((checksDone / 4) * 100),
      csi_status: watch?.csi_status || "",
      papers_status: watch?.papers_status || "",
      customer: row.customer_name || "-",
      route: row.route === "ericdocs" ? "Eric Docs" : "CSI",
      step: Math.min(row.current_step + 1, row.steps.length),
      steps: row.steps.length,
      label: row.complete ? "Done" : step?.label || "",
      status: String(status).slice(0, 200),
      stuck: String(stuck).slice(0, 200),
      papers_sent: watch?.papers_sent === true,
      complete: row.complete === true,
    };
  });
  return {
    date,
    day: {
      complete: day.complete === true,
      step: day.complete ? day.steps.length : Math.min(day.current_step + 1, day.steps.length),
      steps: day.steps.length,
      label: dayStep ? dayStep.label : "Day steps done",
      missing: (dayStep?.missing || []).slice(0, 3),
    },
    rows,
    total: rows.length,
    // Ready = all four green; the % over every shipment's four checks.
    ready: rows.filter((row) => row.pct === 100).length,
    pct: rows.length ? Math.round(rows.reduce((sum, row) => sum + row.pct, 0) / rows.length) : 0,
    done: rows.filter((row) => row.complete).length,
    stuck: rows.filter((row) => row.stuck).length,
    updated_at: new Date().toISOString(),
  };
}

// JSON with its keys sorted (Postgres keeps jsonb keys in its own order).
function stableJson(value) {
  if (Array.isArray(value)) return `[${value.map(stableJson).join(",")}]`;
  if (value && typeof value === "object") return `{${Object.keys(value).sort().map((key) => `${JSON.stringify(key)}:${stableJson(value[key])}`).join(",")}}`;
  return JSON.stringify(value ?? null);
}

export async function syncStepPlanTasks({ today = "", dates: wantedDates, buildPlan }) {
  let dates = wantedDates || [];
  if (!isDatabaseEnabled()) return { made: 0, closed: 0, updated: 0 };
  const config = await readTaskConfig();
  // The older per-step tasks (one per step): replaced by the one per day.
  const older = await dbQuery("SELECT * FROM app_tasks WHERE source = 'step_plan' AND status IN ('open', 'in_progress') AND source_key LIKE 'step_plan:%:%'");
  await closeTasks(older.rows, "system", "Replaced by one step plan task per day.");
  // A day ahead is not tracked (made once before): it closes now, and on
  // its own day it opens again by itself (closed by "system").
  if (today) {
    const ahead = await dbQuery(
      "SELECT * FROM app_tasks WHERE source = 'step_plan' AND status IN ('open', 'in_progress') AND source_key ~ '^step_plan:[0-9]{4}-[0-9]{2}-[0-9]{2}$' AND substr(source_key, 11) > $1",
      [today],
    );
    await closeTasks(ahead.rows, "system", "Not today yet -- it opens again on its own day.");
    older.rows.push(...ahead.rows);
    dates = dates.filter((date) => date <= today);
  }
  if (!config.step_plan.enabled || !config.step_plan.assign_to) return { made: 0, closed: older.rows.length, updated: 0 };
  const assignee = parseAssignee(config.step_plan.assign_to);
  // Earlier days whose task is still open are followed until they are green.
  const stillOpen = await dbQuery(
    "SELECT substr(source_key, 11) AS day FROM app_tasks WHERE source = 'step_plan' AND status IN ('open', 'in_progress') AND source_key ~ '^step_plan:[0-9]{4}-[0-9]{2}-[0-9]{2}$' AND created_at > now() - interval '14 days'",
  );
  dates = [...new Set([...(dates || []), ...stillOpen.rows.map((row) => row.day)])];
  let made = 0;
  let closed = older.rows.length;
  let updated = 0;
  for (const date of dates || []) {
    const built = await buildPlan(date);
    const plan = built?.plan;
    if (!plan) continue;
    // No shipments that day (any more): no task -- an open one closes.
    if (!(plan.rows || []).length) {
      const gone = (await dbQuery("SELECT * FROM app_tasks WHERE source_key = $1 AND status IN ('open', 'in_progress')", [`step_plan:${date}`])).rows;
      await closeTasks(gone, "system", "No shipments for this day in the step plan any more.");
      closed += gone.length;
      continue;
    }
    const progress = stepPlanProgress(date, plan, built.monitor);
    // Done when every shipment is green in the Pipeline Monitor (the step
    // plan's own steps are only shown inside the task).
    const complete = progress.total > 0 && progress.ready === progress.total;
    const title = `Step plan ${date} · ${progress.ready}/${progress.total} shipment(s) ready${progress.stuck ? ` · ${progress.stuck} stuck` : ""}`;
    const sourceKey = `step_plan:${date}`;
    const link = { page: "pdstepplan", date, progress };
    const existing = (await dbQuery("SELECT * FROM app_tasks WHERE source_key = $1", [sourceKey])).rows[0];
    if (!existing) {
      if (complete) continue;
      const row = await createSystemTask({ source: "step_plan", source_key: sourceKey, title, body: "", ...assignee, due_date: date, link });
      if (row) made += 1;
      continue;
    }
    // Keep the progress up to date (the history only gets the open / close).
    const before = stableJson({ title: existing.title, progress: { ...(existing.link?.progress || {}), updated_at: "" } });
    const after = stableJson({ title, progress: { ...progress, updated_at: "" } });
    if (before !== after) {
      await dbQuery("UPDATE app_tasks SET title = $2, link = $3, updated_at = now() WHERE id = $1", [existing.id, title, JSON.stringify(link)]);
      updated += 1;
    }
    if (complete && OPEN_STATUSES.includes(existing.status)) {
      await closeTasks([existing], "system", "The day and every shipment are complete in the step plan.");
      closed += 1;
    } else if (!complete && existing.status === "done" && existing.done_by === "system") {
      await dbQuery("UPDATE app_tasks SET status = 'open', done_by = NULL, done_at = NULL, updated_at = now() WHERE id = $1", [existing.id]);
      await logTaskEvent(existing.id, "system", "status", "done", "open");
      await logTaskEvent(existing.id, "system", "note", "", "Not complete any more (a shipment was added or changed).");
    }
  }
  return { made, closed, updated };
}

// The day's "Facturen Finance Audit" task (workdays): how many facturen are
// still missing, with a Run button for the recorder action on the PC.
export async function syncFacturenTask({ date, withRef, missingRef }) {
  if (!isDatabaseEnabled()) return null;
  const config = await readTaskConfig();
  if (!config.facturen.enabled || !config.facturen.assign_to) return null;
  if ([0, 6].includes(new Date(`${date}T12:00:00Z`).getUTCDay())) return null;
  const sourceKey = `recorder:facturen:${date}`;
  const { title, link } = facturenTaskTitle(date, withRef, missingRef, config.facturen.action);
  const existing = (await dbQuery("SELECT * FROM app_tasks WHERE source_key = $1", [sourceKey])).rows[0];
  if (!existing) {
    if (!withRef) return null;
    return createSystemTask({
      source: "recorder",
      source_key: sourceKey,
      title,
      body: "Fetches every factuur Finance Audit still misses (last 60 days) from Uniware, one by one, and sends them to Finance Audit. Press \u25b6 Run on this PC (the recorder dashboard open, Uniware on screen, hands off the keyboard until it's done). Not ready yet in Uniware = asked again tomorrow.",
      ...parseAssignee(config.facturen.assign_to),
      due_date: date,
      link,
    });
  }
  // The count follows what's really missing, also on a task marked done.
  if (existing.title !== title) {
    await dbQuery("UPDATE app_tasks SET title = $2, link = $3, updated_at = now() WHERE id = $1", [existing.id, title, JSON.stringify(link)]);
  }
  return existing;
}

function facturenTaskTitle(date, withRef, missingRef, action) {
  const total = withRef + missingRef;
  return {
    title: `Facturen Finance Audit ${date} · ${total} missing${missingRef ? ` (${missingRef} without Uniware ref)` : ""}`,
    link: { page: "ukdocs", recorder_action: action, missing: total, with_ref: withRef, missing_ref: missingRef },
  };
}

// A factuur number typed on more than one sending (one is wrong): one task
// per number for the team, closed by itself once the number is on one
// sending only.
export async function syncInvoiceDuplicateTasks({ duplicates }) {
  if (!isDatabaseEnabled()) return { made: 0, closed: 0 };
  const config = await readTaskConfig();
  const keep = new Set();
  let made = 0;
  if (config.invoice_duplicates.enabled) {
    for (const duplicate of duplicates || []) {
      const sourceKey = `invoice_dup:${duplicate.invoice}`;
      keep.add(sourceKey);
      const where = duplicate.sendings.map((owner) => `${owner.customer || "?"} ${owner.date} (${owner.category})`);
      const row = await createSystemTask({
        source: "invoice_dup",
        source_key: sourceKey,
        title: `Factuur ${duplicate.invoice} is on ${duplicate.sendings.length} sendings -- check and fix`,
        body: `The same factuur number is on:\n${where.map((line) => `- ${line}`).join("\n")}\nOne of them is wrong: Finance Audit UKdocs, right-click the wrong sending > Edit invoice numbers. A factuur PDF of the old number is removed then; the next factuur fetch gets the right one.`,
        assignee_team: config.invoice_duplicates.team,
        priority: "high",
        link: { page: "ukdocs", invoice: duplicate.invoice },
      });
      if (row) made += 1;
    }
  }
  const open = await dbQuery("SELECT * FROM app_tasks WHERE source = 'invoice_dup' AND status IN ('open', 'in_progress')");
  const fixed = open.rows.filter((row) => !keep.has(row.source_key));
  await closeTasks(fixed, "system", "The factuur number is on one sending only now.");
  return { made, closed: fixed.length };
}

// A fetch run sent its facturen: the day's task shows the new count and the
// run's note; it's done only when the run got through every factuur it could
// try (one that stopped halfway leaves the task open for the next run).
export async function closeFacturenTask(date, by, note, counts = null) {
  if (!isDatabaseEnabled()) return;
  const row = (await dbQuery("SELECT * FROM app_tasks WHERE source_key = $1", [`recorder:facturen:${date}`])).rows[0];
  if (!row) return;
  if (counts) {
    const config = await readTaskConfig();
    const { title, link } = facturenTaskTitle(date, counts.withRef, counts.missingRef, config.facturen.action);
    if (row.title !== title) {
      await dbQuery("UPDATE app_tasks SET title = $2, link = $3, updated_at = now() WHERE id = $1", [row.id, title, JSON.stringify(link)]);
    }
  }
  if (counts && !counts.finished) {
    if (note) await logTaskEvent(row.id, by, "note", "", note);
    return;
  }
  if (OPEN_STATUSES.includes(row.status)) await closeTasks([row], by, note);
  else if (note) await logTaskEvent(row.id, by, "note", "", note);
}

// ---------------------------------------------------------------------------
// HTTP

export function createAppTasksRoutes({ readUsers, writeUsers, publicUser, sendJson, readRequestJson, stepCatalog, runStepPlanSync }) {
  const prefixes = ["/api/tasks", "/api/chat/", "/api/widget/", "/api/me/settings", "/api/users/directory", "/api/app-teams", "/api/users-teams/"];

  return async function handleAppTasksRoute(req, res, url, requestUser) {
    const pathname = url.pathname;
    if (!prefixes.some((prefix) => pathname === prefix || pathname.startsWith(prefix))) return false;
    try {
      // Own widget settings (in the users file).
      if (pathname === "/api/me/settings") {
        if (req.method === "GET") {
          const users = await readUsers();
          const me = users.find((user) => user.username === requestUser.username);
          sendJson(res, 200, { settings: normalizeUserWidgetSettings(me?.settings), teams: normalizeUserTeams(me?.teams) });
          return true;
        }
        if (req.method === "PUT") {
          const body = await readRequestJson(req);
          const users = await readUsers();
          const index = users.findIndex((user) => user.username === requestUser.username);
          if (index < 0) {
            sendJson(res, 404, { error: "User not found" });
            return true;
          }
          const current = normalizeUserWidgetSettings(users[index].settings);
          const next = normalizeUserWidgetSettings({
            ...current,
            ...(typeof body?.use_widget === "boolean" ? { use_widget: body.use_widget } : {}),
            ...(body?.widget_pos !== undefined ? { widget_pos: body.widget_pos } : {}),
          });
          users[index] = { ...users[index], settings: { ...(users[index].settings || {}), ...next } };
          await writeUsers(users);
          sendJson(res, 200, { user: publicUser(users[index]) });
          return true;
        }
      }

      // Every user's name and teams, for the "for" picker and the DMs.
      if (pathname === "/api/users/directory" && req.method === "GET") {
        const users = await readUsers();
        const config = await readTaskConfig();
        sendJson(res, 200, {
          users: users.map((user) => ({ username: user.username, teams: normalizeUserTeams(user.teams) })).sort((left, right) => left.username.localeCompare(right.username)),
          teams: config.teams,
        });
        return true;
      }

      // Teams + automatic task rules (admin changes them).
      if (pathname === "/api/app-teams") {
        if (req.method === "GET") {
          sendJson(res, 200, { config: await readTaskConfig(), step_catalog: stepCatalog() });
          return true;
        }
        if (req.method === "PUT") {
          if (requestUser.role !== "admin") {
            sendJson(res, 403, { error: "Only an admin changes the teams and task rules." });
            return true;
          }
          const body = await readRequestJson(req, 256 * 1024);
          const current = await readTaskConfig();
          const saved = await writeTaskConfig({ ...current, ...(body?.config || {}) }, requestUser.username);
          sendJson(res, 200, { config: saved, step_catalog: stepCatalog() });
          return true;
        }
      }
      if (pathname === "/api/app-teams/run-step-plan" && req.method === "POST") {
        if (requestUser.role !== "admin") {
          sendJson(res, 403, { error: "Admin only" });
          return true;
        }
        sendJson(res, 200, { result: await runStepPlanSync() });
        return true;
      }

      // A user's teams (admin).
      if (pathname.startsWith("/api/users-teams/") && req.method === "PUT") {
        if (requestUser.role !== "admin") {
          sendJson(res, 403, { error: "Admin only" });
          return true;
        }
        const username = decodeURIComponent(pathname.slice("/api/users-teams/".length));
        const body = await readRequestJson(req);
        const users = await readUsers();
        const index = users.findIndex((user) => user.username === username);
        if (index < 0) {
          sendJson(res, 404, { error: "User not found" });
          return true;
        }
        users[index] = { ...users[index], teams: normalizeUserTeams(body?.teams) };
        if (typeof body?.use_widget === "boolean") {
          users[index].settings = { ...(users[index].settings || {}), ...normalizeUserWidgetSettings({ ...(users[index].settings || {}), use_widget: body.use_widget }) };
        }
        await writeUsers(users);
        sendJson(res, 200, { user: publicUser(users[index]) });
        return true;
      }

      if (!isDatabaseEnabled()) {
        sendJson(res, 409, { error: "Tasks and chat need the database." });
        return true;
      }
      const users = await readUsers();
      const meStored = users.find((user) => user.username === requestUser.username);
      const user = { ...requestUser, teams: normalizeUserTeams(meStored?.teams) };

      // One cheap call for the widget's badges.
      if (pathname === "/api/widget/pulse" && req.method === "GET") {
        const config = await readTaskConfig();
        const counts = await dbQuery(
          `SELECT count(*) AS open,
                  max(id) FILTER (WHERE created_by <> $1) AS newest_id
           FROM app_tasks
           WHERE status IN ('open', 'in_progress') AND (assignee_user = $1 OR (assignee_user IS NULL AND assignee_team = ANY($2::text[])))`,
          [user.username, user.teams],
        );
        const newestId = counts.rows[0]?.newest_id;
        const newest = newestId ? taskRow((await dbQuery("SELECT * FROM app_tasks WHERE id = $1", [newestId])).rows[0]) : null;
        const channels = await listChannels(user, config);
        const unread = channels.reduce((sum, channel) => sum + channel.unread, 0);
        const newestUnread = channels.filter((channel) => channel.unread && channel.last && channel.last.from !== user.username)
          .sort((left, right) => Number(right.last_id) - Number(left.last_id))[0] || null;
        sendJson(res, 200, {
          open_tasks: Number(counts.rows[0]?.open) || 0,
          unread,
          newest_task: newest ? { id: newest.id, title: newest.title, created_by: newest.created_by } : null,
          newest_message: newestUnread ? { id: newestUnread.last_id, channel: newestUnread.channel, title: newestUnread.title, from: newestUnread.last.from, body: newestUnread.last.body } : null,
        });
        return true;
      }

      // ---- tasks ----
      if (pathname === "/api/tasks" && req.method === "GET") {
        const scope = String(url.searchParams.get("scope") || "mine");
        sendJson(res, 200, await listTasks(user, scope));
        return true;
      }
      if (pathname === "/api/tasks" && req.method === "POST") {
        const body = await readRequestJson(req, 64 * 1024);
        const title = String(body?.title || "").trim();
        if (!title) {
          sendJson(res, 400, { error: "A task needs a title." });
          return true;
        }
        let assignee = parseAssignee(body?.assignee);
        if (!assignee.assignee_user && !assignee.assignee_team) assignee = { assignee_user: user.username, assignee_team: null };
        if (assignee.assignee_user && !users.some((item) => item.username === assignee.assignee_user)) {
          sendJson(res, 400, { error: `Unknown user ${assignee.assignee_user}` });
          return true;
        }
        const row = await insertTask({
          title,
          body: body?.body,
          ...assignee,
          created_by: user.username,
          priority: body?.priority,
          due_date: body?.due_date,
          repeat: body?.repeat,
          source: "manual",
          link: body?.link && typeof body.link === "object" ? body.link : {},
        }, user.username);
        sendJson(res, 200, { task: taskRow(row) });
        return true;
      }
      const taskMatch = pathname.match(/^\/api\/tasks\/(\d+)(\/events)?$/);
      if (taskMatch) {
        const task = await getTask(taskMatch[1]);
        if (!task || !canSeeTask(task, user)) {
          sendJson(res, 404, { error: "Task not found" });
          return true;
        }
        if (req.method === "GET") {
          const events = await dbQuery("SELECT at, by_user, action, from_value, to_value FROM app_task_events WHERE task_id = $1 ORDER BY id", [task.id]);
          sendJson(res, 200, { task: taskRow(task), events: events.rows });
          return true;
        }
        if (req.method === "PATCH" && !taskMatch[2]) {
          if (!canChangeTask(task, user)) {
            sendJson(res, 403, { error: "Not your task" });
            return true;
          }
          const body = await readRequestJson(req, 64 * 1024);
          const sets = [];
          const params = [task.id];
          const events = [];
          const set = (column, value, action, from, to) => {
            params.push(value);
            sets.push(`${column} = $${params.length}`);
            if (action) events.push([action, from, to]);
          };
          if (body?.status !== undefined && TASK_STATUSES.includes(body.status) && body.status !== task.status) {
            set("status", body.status, "status", task.status, body.status);
            if (body.status === "done") {
              set("done_by", user.username);
              sets.push("done_at = now()");
            } else {
              sets.push("done_by = NULL", "done_at = NULL");
            }
          }
          if (body?.take === true && task.assignee_user !== user.username) {
            set("assignee_user", user.username, "assignee", task.assignee_user ? `user:${task.assignee_user}` : `team:${task.assignee_team || ""}`, `user:${user.username}`);
          } else if (body?.assignee !== undefined) {
            const next = parseAssignee(body.assignee);
            if (next.assignee_user && !users.some((item) => item.username === next.assignee_user)) {
              sendJson(res, 400, { error: `Unknown user ${next.assignee_user}` });
              return true;
            }
            if (next.assignee_user || next.assignee_team) {
              set("assignee_user", next.assignee_user, "assignee", task.assignee_user ? `user:${task.assignee_user}` : `team:${task.assignee_team || ""}`, body.assignee);
              set("assignee_team", next.assignee_team);
            }
          }
          if (body?.priority !== undefined && TASK_PRIORITIES.includes(body.priority) && body.priority !== task.priority) set("priority", body.priority, "priority", task.priority, body.priority);
          if (body?.due_date !== undefined) {
            const due = isoDay(body.due_date) || null;
            const before = taskRow(task).due_date;
            if ((due || "") !== before) set("due_date", due, "due_date", before, due || "");
          }
          if (body?.repeat !== undefined && TASK_REPEATS.includes(body.repeat) && body.repeat !== task.repeat) set("repeat", body.repeat, "repeat", task.repeat, body.repeat);
          if (typeof body?.title === "string" && body.title.trim() && body.title.trim() !== task.title) set("title", body.title.trim().slice(0, 300), "title", task.title, body.title.trim());
          if (typeof body?.body === "string" && body.body !== task.body) set("body", body.body.slice(0, 8000), "body", "", "");
          if (!sets.length) {
            sendJson(res, 200, { task: taskRow(task) });
            return true;
          }
          sets.push("updated_at = now()");
          const updated = (await dbQuery(`UPDATE app_tasks SET ${sets.join(", ")} WHERE id = $1 RETURNING *`, params)).rows[0];
          for (const [action, from, to] of events) await logTaskEvent(task.id, user.username, action, from, to);
          // A repeating task done: its next one.
          let next = null;
          if (body?.status === "done" && task.status !== "done" && TASK_REPEATS.includes(updated.repeat) && updated.repeat !== "none") {
            next = await insertTask({
              title: updated.title,
              body: updated.body,
              assignee_user: updated.assignee_team ? null : updated.assignee_user,
              assignee_team: updated.assignee_team,
              created_by: updated.created_by,
              priority: updated.priority,
              due_date: nextRepeatDay(updated.repeat, taskRow(updated).due_date || localToday()),
              repeat: updated.repeat,
              source: updated.source === "manual" ? "manual" : updated.source,
              link: updated.link,
            }, "system");
          }
          sendJson(res, 200, { task: taskRow(updated), next: taskRow(next) });
          return true;
        }
      }

      // ---- chat ----
      if (pathname === "/api/chat/channels" && req.method === "GET") {
        const config = await readTaskConfig();
        sendJson(res, 200, { channels: await listChannels(user, config) });
        return true;
      }
      if (pathname === "/api/chat/messages" && req.method === "GET") {
        let channel = String(url.searchParams.get("channel") || "");
        const withUser = String(url.searchParams.get("with") || "").trim();
        if (withUser) channel = dmChannel(user.username, withUser);
        if (!(await canUseChannel(channel, user, users))) {
          sendJson(res, 403, { error: "Not your conversation" });
          return true;
        }
        const before = /^\d+$/.test(url.searchParams.get("before_id") || "") ? url.searchParams.get("before_id") : null;
        const after = /^\d+$/.test(url.searchParams.get("after_id") || "") ? url.searchParams.get("after_id") : null;
        const result = after
          ? await dbQuery("SELECT * FROM app_chat_messages WHERE channel = $1 AND id > $2 ORDER BY id LIMIT 200", [channel, after])
          : await dbQuery(
            `SELECT * FROM (SELECT * FROM app_chat_messages WHERE channel = $1 ${before ? "AND id < $2" : ""} ORDER BY id DESC LIMIT 50) last ORDER BY id`,
            before ? [channel, before] : [channel],
          );
        const read = await dbQuery("SELECT last_read_id FROM app_chat_reads WHERE user_name = $1 AND channel = $2", [user.username, channel]);
        sendJson(res, 200, { channel, messages: result.rows.map(messageRow), last_read_id: String(read.rows[0]?.last_read_id || 0), has_more: !after && result.rows.length === 50 });
        return true;
      }
      if (pathname === "/api/chat/messages" && req.method === "POST") {
        const body = await readRequestJson(req, 64 * 1024);
        let channel = String(body?.channel || "");
        if (body?.to_user) channel = dmChannel(user.username, String(body.to_user).trim());
        const text = String(body?.body || "").trim().slice(0, 4000);
        if (!text) {
          sendJson(res, 400, { error: "Empty message" });
          return true;
        }
        if (!(await canUseChannel(channel, user, users))) {
          sendJson(res, 403, { error: "Not your conversation" });
          return true;
        }
        const row = (await dbQuery("INSERT INTO app_chat_messages (channel, from_user, body) VALUES ($1, $2, $3) RETURNING *", [channel, user.username, text])).rows[0];
        await dbQuery(
          `INSERT INTO app_chat_reads (user_name, channel, last_read_id) VALUES ($1, $2, $3)
           ON CONFLICT (user_name, channel) DO UPDATE SET last_read_id = GREATEST(app_chat_reads.last_read_id, EXCLUDED.last_read_id)`,
          [user.username, channel, row.id],
        );
        sendJson(res, 200, { message: messageRow(row) });
        return true;
      }
      if (pathname === "/api/chat/read" && req.method === "POST") {
        const body = await readRequestJson(req);
        const channel = String(body?.channel || "");
        const lastId = /^\d+$/.test(String(body?.last_id || "")) ? String(body.last_id) : "0";
        if (!(await canUseChannel(channel, user, users))) {
          sendJson(res, 403, { error: "Not your conversation" });
          return true;
        }
        await dbQuery(
          `INSERT INTO app_chat_reads (user_name, channel, last_read_id) VALUES ($1, $2, $3)
           ON CONFLICT (user_name, channel) DO UPDATE SET last_read_id = GREATEST(app_chat_reads.last_read_id, EXCLUDED.last_read_id)`,
          [user.username, channel, lastId],
        );
        sendJson(res, 200, { ok: true });
        return true;
      }
      const deleteMatch = pathname.match(/^\/api\/chat\/messages\/(\d+)$/);
      if (deleteMatch && req.method === "DELETE") {
        const row = (await dbQuery("SELECT * FROM app_chat_messages WHERE id = $1", [deleteMatch[1]])).rows[0];
        if (!row || (row.from_user !== user.username && user.role !== "admin")) {
          sendJson(res, 404, { error: "Message not found" });
          return true;
        }
        // Kept (soft delete): the text stays in the database, only hidden.
        await dbQuery("UPDATE app_chat_messages SET deleted = true, edited_at = now() WHERE id = $1", [row.id]);
        sendJson(res, 200, { ok: true });
        return true;
      }

      sendJson(res, 404, { error: "Not found" });
      return true;
    } catch (error) {
      sendJson(res, 400, { error: error instanceof Error ? error.message : String(error) });
      return true;
    }
  };
}
