require('dotenv').config();
const express = require('express');
const fs = require('fs');
const path = require('path');
const teams = require('./teams');

const app = express();
const PORT = process.env.PORT || 3000;

const METABASE_URL = (process.env.METABASE_URL || 'https://mb.syncfuze.com').replace(/\/$/, '');
const METABASE_USER = process.env.METABASE_USER;
const METABASE_PASS = process.env.METABASE_PASS;

// The databases shown on the dashboard are managed from the UI and saved here.
// On first run the list is seeded from METABASE_DATABASE_NAME (comma-separated).
const DATABASES_FILE = path.join(__dirname, 'databases.json');

function loadDatabaseNames() {
  try {
    const saved = JSON.parse(fs.readFileSync(DATABASES_FILE, 'utf8'));
    if (Array.isArray(saved)) return saved.filter(n => typeof n === 'string' && n.trim());
  } catch (err) {
    if (err.code !== 'ENOENT') console.error('Could not read databases.json:', err.message);
  }
  return (process.env.METABASE_DATABASE_NAME || '')
    .split(',')
    .map(n => n.trim())
    .filter(Boolean);
}

let databaseNames = loadDatabaseNames();

function saveDatabaseNames() {
  fs.writeFileSync(DATABASES_FILE, JSON.stringify(databaseNames, null, 2) + '\n');
}

// Databases of the same customer project share a base name:
// ajg-SDB, ajg2-SDB, ajg3-SDB, ajg4 -> AJG; horbergmsg, horbergemail-new -> HORBERG
function projectOf(database) {
  const original = database.trim().toLowerCase();
  let n = original, prev;
  do {
    prev = n;
    n = n
      .replace(/[-_ ]?(sdb|new)$/, '')
      .replace(/[-_ ]?\d+$/, '')
      .replace(/(msg|message|messages|email|emails|mail)$/, '')
      .replace(/[-_ ]+$/, '');
  } while (n !== prev && n);
  return (n || original).toUpperCase();
}

// Per-project Teams settings: { "AJG": { webhookUrl, hourly, lastSentAt, lastError } }.
// Holds webhook URLs, which act as secrets, so this file is git-ignored.
const PROJECTS_FILE = path.join(__dirname, 'projects.json');

function loadProjectSettings() {
  try {
    const saved = JSON.parse(fs.readFileSync(PROJECTS_FILE, 'utf8'));
    if (saved && typeof saved === 'object') return saved;
  } catch (err) {
    if (err.code !== 'ENOENT') console.error('Could not read projects.json:', err.message);
  }
  return {};
}

let projectSettings = loadProjectSettings();

function saveProjectSettings() {
  fs.writeFileSync(PROJECTS_FILE, JSON.stringify(projectSettings, null, 2) + '\n');
}

function publicProjectSettings(project) {
  const s = projectSettings[project] || {};
  return {
    project,
    webhookConfigured: Boolean(s.webhookUrl),
    webhookPreview: teams.maskWebhookUrl(s.webhookUrl),
    hourly: Boolean(s.hourly && s.webhookUrl),
    lastSentAt: s.lastSentAt || null,
    lastError: s.lastError || null
  };
}

// Workspaces owned by these email domains (internal test accounts) are left out
// of every count, along with their files. Comma-separated.
const EXCLUDED_OWNER_DOMAINS = (process.env.EXCLUDED_OWNER_DOMAINS ?? 'cloudfuze.com')
  .split(',')
  .map(d => d.trim().replace(/^@/, ''))
  .filter(Boolean);
const escapeRegex = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const EXCLUDED_OWNER_REGEX = EXCLUDED_OWNER_DOMAINS.length
  ? { $regex: `@(${EXCLUDED_OWNER_DOMAINS.map(escapeRegex).join('|')})$`, $options: 'i' }
  : null;

let sessionToken = null;
let databaseList = null;

async function metabaseFetch(pathPart, options = {}) {
  const headers = { 'Content-Type': 'application/json', ...(options.headers || {}) };
  if (sessionToken) headers['X-Metabase-Session'] = sessionToken;
  const res = await fetch(`${METABASE_URL}${pathPart}`, { ...options, headers });
  if (!res.ok) {
    const body = await res.text();
    let detail = body;
    try { detail = JSON.parse(body).error || body; } catch {}
    const err = new Error(`Metabase ${res.status} on ${pathPart}: ${detail}`);
    err.status = res.status;
    throw err;
  }
  return res.json();
}

async function login() {
  if (!METABASE_USER || !METABASE_PASS) {
    throw new Error('METABASE_USER and METABASE_PASS must be set in .env');
  }
  const res = await fetch(`${METABASE_URL}/api/session`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ username: METABASE_USER, password: METABASE_PASS })
  });
  if (!res.ok) {
    throw new Error(`Login failed (${res.status}): ${await res.text()}`);
  }
  const data = await res.json();
  sessionToken = data.id;
  console.log('Logged into Metabase as', METABASE_USER);
}

async function withRelogin(fn) {
  if (!sessionToken) await login();
  try {
    return await fn();
  } catch (err) {
    if (err.status === 401 || err.status === 403) {
      sessionToken = null;
      await login();
      return await fn();
    }
    throw err;
  }
}

async function getMetabaseDatabases({ refresh = false } = {}) {
  if (!databaseList || refresh) {
    const data = await withRelogin(() => metabaseFetch('/api/database'));
    databaseList = Array.isArray(data) ? data : (data.data || []);
  }
  return databaseList;
}

const sameName = (a, b) => a.trim().toLowerCase() === b.trim().toLowerCase();

async function findDatabaseId(name) {
  const db = (await getMetabaseDatabases()).find(d => d.name && sameName(d.name, name));
  if (!db) {
    // Refresh the cached list next time in case the database was added since
    databaseList = null;
    throw new Error(`Database "${name}" not found in Metabase`);
  }
  return db.id;
}

// Cached per database ID so metadata is only fetched once per server run.
const dbTypeCache = new Map();
const CONTENT_COLLECTIONS = ['FilefolderInfo', 'FolderMetadataInfo', 'CollaborationDetails', 'Hyperlinks'];

async function getDatabaseType(databaseId) {
  if (dbTypeCache.has(databaseId)) return dbTypeCache.get(databaseId);
  try {
    const data = await withRelogin(() => metabaseFetch(`/api/database/${databaseId}/metadata`));
    const tables = new Set((data.tables || []).map(t => t.name.toLowerCase().replace(/[_\s]/g, '')));
    const isContent = CONTENT_COLLECTIONS.some(c => tables.has(c.toLowerCase()));
    const type = isContent ? 'content' : 'message';
    dbTypeCache.set(databaseId, type);
    return type;
  } catch {
    return 'message';
  }
}

// MongoDB errors that mean the database server is briefly unavailable (failover,
// a replica set member recovering, network blips). These are worth retrying.
const TRANSIENT_MONGO_ERROR = /NotPrimaryOrSecondary|node is recovering|NotWritablePrimary|NotMaster|PrimarySteppedDown|InterruptedDueToReplStateChange|HostUnreachable|SocketException|connection .*(closed|reset)/i;
const RETRY_DELAYS_MS = [1500, 4000];

const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));

// Turns raw Metabase/MongoDB errors into a short message for the dashboard
function friendlyError(err) {
  const msg = String(err?.message || err);
  if (TRANSIENT_MONGO_ERROR.test(msg)) {
    return 'The MongoDB server for this database is temporarily unavailable (it reported "node is recovering" or a similar failover state). ' +
      'This is a database server issue, not the dashboard. Try Refresh in a few minutes.';
  }
  const short = msg.replace(/^Metabase \d+ on \/api\/dataset: /, '');
  return short.length > 300 ? short.slice(0, 300) + '…' : short;
}

async function runMongoQuery(databaseId, collection, pipeline) {
  // The Metabase databases are MongoDB, so queries are aggregation pipelines, not SQL
  const body = {
    database: databaseId,
    type: 'native',
    native: { collection, query: JSON.stringify(pipeline) }
  };

  const send = async () => {
    const result = await withRelogin(() => metabaseFetch('/api/dataset', {
      method: 'POST',
      body: JSON.stringify(body)
    }));
    // Metabase reports query errors with HTTP 202 and an error field
    if (result?.error) throw new Error(result.error);
    return result;
  };

  let result;
  for (let attempt = 0; ; attempt++) {
    try {
      result = await send();
      break;
    } catch (err) {
      if (attempt >= RETRY_DELAYS_MS.length || !TRANSIENT_MONGO_ERROR.test(err.message)) throw err;
      await sleep(RETRY_DELAYS_MS[attempt]);
    }
  }

  const names = (result?.data?.cols || []).map(c => c.name);
  return (result?.data?.rows || []).map(r => Object.fromEntries(names.map((n, i) => [n, r[i]])));
}

// IDs (as strings) of the users whose primaryEmail is not in an excluded domain
async function findIncludedUserIds(databaseId) {
  const rows = await runMongoQuery(databaseId, 'Users', [
    { $match: { primaryEmail: { $not: EXCLUDED_OWNER_REGEX } } },
    { $group: { _id: null, ids: { $push: { $toString: '$_id' } } } }
  ]);
  const ids = rows[0]?.ids || [];
  return typeof ids === 'string' ? JSON.parse(ids) : ids;
}

// Counts by processStatus, as [status, count] rows
async function runProcessStatusQuery(databaseId, collection, match) {
  const pipeline = [
    ...(match ? [{ $match: match }] : []),
    { $group: { _id: '$processStatus', Count: { $sum: 1 } } },
    { $project: { _id: 0, ProcessStatus: '$_id', Count: 1 } },
    { $sort: { ProcessStatus: 1 } }
  ];
  const rows = await runMongoQuery(databaseId, collection, pipeline);
  return rows.map(r => [r.ProcessStatus, r.Count]);
}

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Loads all 5 content collections: MoveWorkspace first, then the 4 file/folder ones.
async function loadContentDatabaseStatus(name, databaseId) {
  async function queryContent(collection) {
    try {
      return { collection, rows: await runProcessStatusQuery(databaseId, collection, null) };
    } catch (err) {
      console.error(`Content query failed for ${name}.${collection}:`, err.message);
      return { collection, error: friendlyError(err) };
    }
  }

  const [moveWorkspace, filefolderInfo, folderMetadataInfo, collaborationDetails, hyperlinks] =
    await Promise.all([
      queryContent('MoveWorkspace'),
      queryContent('FilefolderInfo'),
      queryContent('FolderMetadataInfo'),
      queryContent('CollaborationDetails'),
      queryContent('Hyperlinks'),
    ]);

  // If every query failed with the same error the DB server is likely down
  const all = [moveWorkspace, filefolderInfo, folderMetadataInfo, collaborationDetails, hyperlinks];
  const errors = all.map(c => c.error);
  if (errors.every(e => e) && new Set(errors).size === 1) {
    return { database: name, dbType: 'content', error: errors[0] };
  }

  let combinations = [];
  try {
    combinations = await findCombinations(databaseId, null, 'MoveWorkspace');
  } catch (err) {
    console.error(`Combination lookup failed for ${name}:`, err.message);
  }

  return {
    database: name,
    dbType: 'content',
    combinations,
    workspace: [moveWorkspace],
    files: [filefolderInfo, folderMetadataInfo, collaborationDetails, hyperlinks]
  };
}

async function loadDatabaseStatus(name) {
  let databaseId;
  try {
    databaseId = await findDatabaseId(name);
  } catch (err) {
    console.error(`Lookup failed for ${name}:`, err.message);
    return { database: name, error: friendlyError(err) };
  }

  const dbType = await getDatabaseType(databaseId);
  if (dbType === 'content') {
    return loadContentDatabaseStatus(name, databaseId);
  }

  const workspaceBase = EXCLUDED_OWNER_REGEX ? { ownerEmailId: { $not: EXCLUDED_OWNER_REGEX } } : null;

  // User IDs fetched once and shared across all three MessageEachFiles queries
  const userIdsPromise = EXCLUDED_OWNER_REGEX ? findIncludedUserIds(databaseId) : Promise.resolve(null);

  async function queryWS(label, dmFilter) {
    try {
      const match = dmFilter !== undefined
        ? (workspaceBase ? { ...workspaceBase, directOrGroupMessage: dmFilter } : { directOrGroupMessage: dmFilter })
        : workspaceBase;
      return { collection: label, rows: await runProcessStatusQuery(databaseId, 'MessageWorkSpace', match) };
    } catch (err) {
      console.error(`Query failed for ${name}.MessageWorkSpace (${label}):`, err.message);
      return { collection: label, error: friendlyError(err) };
    }
  }

  async function queryEF(label, dmFilter) {
    try {
      const ids = await userIdsPromise;
      const base = ids ? { userId: { $in: ids } } : null;
      const match = dmFilter !== undefined
        ? (base ? { ...base, directOrGroupMessage: dmFilter } : { directOrGroupMessage: dmFilter })
        : base;
      return { collection: label, rows: await runProcessStatusQuery(databaseId, 'MessageEachFiles', match) };
    } catch (err) {
      console.error(`Query failed for ${name}.MessageEachFiles (${label}):`, err.message);
      return { collection: label, error: friendlyError(err) };
    }
  }

  const [wsTotal, wsChannels, wsDMs, efTotal, efChannels, efDMs] = await Promise.all([
    queryWS('MessageWorkSpace', undefined),   // all workspaces
    queryWS('Channels',         false),       // directOrGroupMessage = false
    queryWS('Direct Messages',  true),        // directOrGroupMessage = true
    queryEF('MessageEachFiles', undefined),   // all files
    queryEF('Channels',         false),
    queryEF('Direct Messages',  true),
  ]);

  // If the database server itself is down every query fails the same way;
  // show that once for the database instead of once per collection
  const allResults = [wsTotal, wsChannels, wsDMs, efTotal, efChannels, efDMs];
  const errors = allResults.map(c => c.error);
  if (errors.every(e => e) && new Set(errors).size === 1) {
    return { database: name, error: errors[0] };
  }

  let combinations = [];
  try {
    combinations = await findCombinations(databaseId, workspaceBase);
  } catch (err) {
    console.error(`Combination lookup failed for ${name}:`, err.message);
  }

  return {
    database: name,
    combinations,
    workspace: [wsTotal, wsChannels, wsDMs],
    files:     [efTotal, efChannels, efDMs]
  };
}

// Source -> destination cloud pairs, most items first.
// collection is 'MessageWorkSpace' for message projects, 'MoveWorkspace' for content projects.
// Falls back to unfiltered when the owner filter leaves no results.
async function findCombinations(databaseId, workspaceMatch, collection = 'MessageWorkSpace') {
  const pipeline = match => [
    ...(match ? [{ $match: match }] : []),
    { $group: { _id: { from: '$fromCloudName', to: '$toCloudName' }, Count: { $sum: 1 } } },
    { $project: { _id: 0, from: '$_id.from', to: '$_id.to', Count: 1 } },
    { $sort: { Count: -1 } }
  ];
  let rows = await runMongoQuery(databaseId, collection, pipeline(workspaceMatch));
  if (!rows.length && workspaceMatch) rows = await runMongoQuery(databaseId, collection, pipeline(null));
  return rows
    .filter(r => r.from || r.to)
    .map(r => ({ from: r.from, to: r.to, count: r.Count }));
}

// Streaming endpoint: sends each database result as an SSE event as it finishes,
// so the UI can render progressively instead of waiting for all databases.
app.get('/api/process-status/stream', async (req, res) => {
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.flushHeaders();

  const names = databaseNames;
  const send = data => res.write(`data: ${JSON.stringify(data)}\n\n`);

  send({ type: 'init', databases: names });

  await Promise.all(names.map(async name => {
    try {
      const result = await loadDatabaseStatus(name);
      send({ type: 'result', ...result, project: projectOf(name) });
    } catch (err) {
      send({ type: 'result', database: name, error: err.message, project: projectOf(name) });
    }
  }));

  const projects = Object.fromEntries(
    [...new Set(names.map(n => projectOf(n)))].map(p => [p, publicProjectSettings(p)])
  );
  send({ type: 'done', projects, excludedOwnerDomains: EXCLUDED_OWNER_DOMAINS, fetchedAt: new Date().toISOString() });
  res.end();
});

// All dashboard databases, or just one with ?database=<name>
app.get('/api/process-status', async (req, res) => {
  const names = req.query.database ? [String(req.query.database)] : databaseNames;
  const results = await Promise.all(names.map(async name => ({
    ...(await loadDatabaseStatus(name)),
    project: projectOf(name)
  })));
  const projects = Object.fromEntries(
    [...new Set(results.map(r => r.project))].map(p => [p, publicProjectSettings(p)])
  );
  res.json({ results, projects, excludedOwnerDomains: EXCLUDED_OWNER_DOMAINS, fetchedAt: new Date().toISOString() });
});

// ---- Teams: per-project status posts ----

const projectDatabases = project => databaseNames.filter(n => projectOf(n) === project);

// Natural sort: compare digit runs numerically, everything else lexicographically.
// Ensures ajg-SDB → ajg2-SDB → ajg3-SDB → ajg4 (hyphen before any digit).
function naturalSort(a, b) {
  const parts = s => s.toLowerCase().split(/(\d+)/);
  const pa = parts(a), pb = parts(b);
  for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
    const sa = pa[i] ?? '', sb = pb[i] ?? '';
    const na = Number(sa), nb = Number(sb);
    if (!isNaN(na) && !isNaN(nb) && sa !== '' && sb !== '') {
      if (na !== nb) return na - nb;
    } else if (sa !== sb) {
      return sa < sb ? -1 : 1;
    }
  }
  return 0;
}

// Loads all results for a project, sorted by database name
async function loadProjectResults(project) {
  const names = projectDatabases(project).sort(naturalSort);
  if (!names.length) throw new Error(`${project} has no databases on the dashboard`);
  return Promise.all(names.map(loadDatabaseStatus));
}

// One card per database so each card stays within Teams' size limit regardless
// of how many databases a project has (AJG has 4; Morris has 1).
async function sendProjectStatus(project) {
  const settings = projectSettings[project];
  if (!settings?.webhookUrl) throw new Error(`No Teams webhook is set up for ${project}`);
  const results = await loadProjectResults(project);
  const opts = { excludedDomains: EXCLUDED_OWNER_DOMAINS };
  try {
    for (const r of results) {
      const card = teams.buildCard(project, [r], opts);
      console.log(`Teams card for ${r.database}: ${JSON.stringify(card).length} bytes`);
      await teams.postToWebhook(settings.webhookUrl, card);
    }
    settings.lastSentAt = new Date().toISOString();
    settings.lastError = null;
    console.log(`Sent ${project} status to Teams (${results.length} card${results.length === 1 ? '' : 's'})`);
  } catch (err) {
    settings.lastError = `${new Date().toISOString()}: ${err.message}`;
    throw err;
  } finally {
    saveProjectSettings();
  }
}

// Preview endpoint uses the first database only
async function buildProjectCard(project) {
  const results = await loadProjectResults(project);
  return teams.buildCard(project, results.slice(0, 1), { excludedDomains: EXCLUDED_OWNER_DOMAINS });
}

app.put('/api/projects/:project/teams', (req, res) => {
  const project = String(req.params.project).toUpperCase();
  const current = projectSettings[project] || {};
  const next = { ...current };
  try {
    if (req.body?.webhookUrl !== undefined) {
      next.webhookUrl = req.body.webhookUrl ? teams.validateWebhookUrl(req.body.webhookUrl) : '';
      if (!next.webhookUrl) next.hourly = false;
    }
    if (req.body?.hourly !== undefined) {
      if (req.body.hourly && !next.webhookUrl) throw new Error('Add the Teams webhook URL before turning on hourly updates');
      next.hourly = Boolean(req.body.hourly);
    }
  } catch (err) {
    return res.status(400).json({ error: err.message });
  }
  projectSettings[project] = next;
  saveProjectSettings();
  res.json(publicProjectSettings(project));
});

app.post('/api/projects/:project/teams/send', async (req, res) => {
  const project = String(req.params.project).toUpperCase();
  try {
    await sendProjectStatus(project);
    res.json(publicProjectSettings(project));
  } catch (err) {
    res.status(502).json({ error: err.message, ...publicProjectSettings(project) });
  }
});

// Returns the card JSON without posting — paste into adaptivecards.io/designer to validate
app.get('/api/projects/:project/teams/card', async (req, res) => {
  const project = String(req.params.project).toUpperCase();
  try {
    const card = await buildProjectCard(project);
    res.json(card);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Posts a tiny test card to verify the webhook URL is still alive
app.post('/api/projects/:project/teams/ping', async (req, res) => {
  const project = String(req.params.project).toUpperCase();
  const settings = projectSettings[project];
  if (!settings?.webhookUrl) return res.status(400).json({ error: 'No webhook URL saved for this project' });
  const pingCard = {
    type: 'message',
    attachments: [{
      contentType: 'application/vnd.microsoft.card.adaptive',
      contentUrl: null,
      content: {
        $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
        type: 'AdaptiveCard',
        version: '1.4',
        msteams: { width: 'Full' },
        body: [{ type: 'TextBlock', text: `Webhook test for ${project} — if you see this in Teams the connection is working.`, wrap: true }]
      }
    }]
  };
  try {
    await teams.postToWebhook(settings.webhookUrl, pingCard);
    res.json({ ok: true, message: 'Ping sent — check Teams for the test message' });
  } catch (err) {
    res.status(502).json({ error: err.message });
  }
});

// Sends every project with hourly updates on, at the top of each hour
async function sendHourlyUpdates() {
  const due = Object.keys(projectSettings).filter(p => projectSettings[p].hourly && projectSettings[p].webhookUrl);
  for (const project of due) {
    try {
      await sendProjectStatus(project);
    } catch (err) {
      console.error(`Hourly Teams update failed for ${project}:`, err.message);
    }
  }
}

function scheduleHourlyUpdates() {
  const now = new Date();
  const next = new Date(now);
  next.setHours(now.getHours() + 1, 0, 0, 0);
  setTimeout(async () => {
    await sendHourlyUpdates();
    scheduleHourlyUpdates();
  }, next - now);
  console.log(`Next hourly Teams update at ${next.toLocaleTimeString()}`);
}

// The dashboard's databases plus every database name Metabase offers, for the picker
app.get('/api/databases', async (req, res) => {
  try {
    const all = await getMetabaseDatabases({ refresh: true });
    res.json({
      selected: databaseNames,
      available: all.map(d => d.name.trim()).sort((a, b) => a.localeCompare(b))
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/databases', async (req, res) => {
  const name = String(req.body?.name || '').trim();
  if (!name) return res.status(400).json({ error: 'Database name is required' });
  try {
    const match = (await getMetabaseDatabases({ refresh: true })).find(d => d.name && sameName(d.name, name));
    if (!match) return res.status(404).json({ error: `Database "${name}" not found in Metabase` });
    const canonical = match.name.trim();
    if (databaseNames.some(n => sameName(n, canonical))) {
      return res.status(409).json({ error: `"${canonical}" is already on the dashboard` });
    }
    databaseNames.push(canonical);
    saveDatabaseNames();
    res.status(201).json({ name: canonical, selected: databaseNames });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.delete('/api/databases/:name', (req, res) => {
  const before = databaseNames.length;
  databaseNames = databaseNames.filter(n => !sameName(n, req.params.name));
  if (databaseNames.length === before) {
    return res.status(404).json({ error: `"${req.params.name}" is not on the dashboard` });
  }
  saveDatabaseNames();
  res.json({ selected: databaseNames });
});

app.listen(PORT, () => {
  console.log(`Dashboard running at http://localhost:${PORT}`);
  scheduleHourlyUpdates();
});
