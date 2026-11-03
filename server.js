require('dotenv').config();
const express = require('express');
const session = require('express-session');
const msal = require('@azure/msal-node');
const fs = require('fs');
const path = require('path');
const teams = require('./teams');

const app = express();
const PORT = process.env.PORT || 3000;

const METABASE_URL = (process.env.METABASE_URL || 'https://mb.syncfuze.com').replace(/\/$/, '');
const METABASE_USER = process.env.METABASE_USER;
const METABASE_PASS = process.env.METABASE_PASS;

// ---- Microsoft SSO ----
const AZURE_CLIENT_ID     = process.env.AZURE_CLIENT_ID;
const AZURE_TENANT_ID     = process.env.AZURE_TENANT_ID || 'common';
const AZURE_CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;
const AZURE_REDIRECT_URI  = process.env.AZURE_REDIRECT_URI || `http://localhost:${PORT}/auth/callback`;
const SESSION_SECRET       = process.env.SESSION_SECRET || 'change-me-please';
const ALLOWED_DOMAIN       = (process.env.ALLOWED_EMAIL_DOMAIN || 'cloudfuze.com').toLowerCase();

const msalApp = AZURE_CLIENT_ID && AZURE_CLIENT_SECRET
  ? new msal.ConfidentialClientApplication({
      auth: {
        clientId: AZURE_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${AZURE_TENANT_ID}`,
        clientSecret: AZURE_CLIENT_SECRET,
      }
    })
  : null;

if (!msalApp) {
  console.warn('AZURE_CLIENT_ID / AZURE_CLIENT_SECRET not set — SSO is disabled, dashboard is open to all.');
}

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

// Per-project Teams settings: { "AJG": { webhookUrl, interval, lastSentAt, lastError } }.
// interval = minutes between auto-posts (0 = off, 30, 60, 120, 180). Legacy boolean `hourly` is treated as interval=60.
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
    interval: s.webhookUrl ? (s.interval ?? (s.hourly ? 60 : 0)) : 0,
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
const dbTypeCache = new Map();

// ---------------------------------------------------------------------------
// Hourly snapshot – stores the previous hour's counts per database/collection
// so the dashboard can show deltas like "+150 last hr".
// ---------------------------------------------------------------------------
const SNAPSHOTS_FILE = path.join(__dirname, 'snapshots.json');
const SNAPSHOT_MIN_AGE_MS = 5 * 60 * 1000;   // don't diff a just-taken snapshot
const SNAPSHOT_ROLL_MS    = 60 * 60 * 1000;   // roll baseline forward every hour

function loadSnapshots() {
  try {
    const saved = JSON.parse(fs.readFileSync(SNAPSHOTS_FILE, 'utf8'));
    if (saved && typeof saved === 'object') return saved;
  } catch (err) {
    if (err.code !== 'ENOENT') console.error('Could not read snapshots.json:', err.message);
  }
  return {};
}

let hourlySnapshots = loadSnapshots();

function saveSnapshots() {
  try {
    fs.writeFileSync(SNAPSHOTS_FILE, JSON.stringify(hourlySnapshots, null, 2) + '\n');
  } catch (err) {
    console.error('Could not save snapshots.json:', err.message);
  }
}

// Attaches delta info to each collection by comparing against the stored
// hourly snapshot, then rolls the snapshot forward when it is ≥ 1 hour old.
function applySnapshots(name, result) {
  if (result.error) return result;

  const now = Date.now();
  const snap = hourlySnapshots[name];
  const allCollections = [...(result.workspace || []), ...(result.files || [])];

  // Build status→count maps for every healthy collection
  const currentCounts = {};
  for (const c of allCollections) {
    if (!c.error && Array.isArray(c.rows)) {
      currentCounts[c.collection] = Object.fromEntries(
        c.rows.map(([s, n]) => [s ?? '(null)', Number(n) || 0])
      );
    }
  }

  // Compute deltas only when snapshot is old enough to be meaningful
  let deltas = null;
  if (snap && snap.timestamp && (now - snap.timestamp) >= SNAPSHOT_MIN_AGE_MS) {
    deltas = {};
    for (const [collName, current] of Object.entries(currentCounts)) {
      const prev = snap.collections?.[collName] || {};
      let totalDelta = 0;
      const statusDeltas = {};
      for (const [s, n] of Object.entries(current)) {
        const d = n - (prev[s] || 0);
        statusDeltas[s] = d;
        totalDelta += d;
      }
      for (const [s, n] of Object.entries(prev)) {
        if (!(s in current)) { statusDeltas[s] = -n; totalDelta -= n; }
      }
      const previousTotal = Object.values(prev).reduce((s, n) => s + n, 0);
      deltas[collName] = { totalDelta, statusDeltas, snapshotAt: snap.timestamp, previousTotal };
    }
  }

  // Roll snapshot forward once it is at least 1 hour old (or brand-new)
  if (!snap || (now - snap.timestamp) >= SNAPSHOT_ROLL_MS) {
    hourlySnapshots[name] = { timestamp: now, collections: currentCounts };
    saveSnapshots();
  }

  if (!deltas) return result;

  const attach = cols => (cols || []).map(c =>
    (!c.error && deltas[c.collection]) ? { ...c, delta: deltas[c.collection] } : c
  );
  return { ...result, workspace: attach(result.workspace), files: attach(result.files) };
}

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

// Detect whether a database is a Messaging, Content, or Email project by inspecting
// its collection names via the Metabase metadata API. Result is cached per database ID.
async function detectDatabaseType(databaseId) {
  if (dbTypeCache.has(databaseId)) return dbTypeCache.get(databaseId);
  let type = 'messaging';
  try {
    const meta = await withRelogin(() => metabaseFetch(`/api/database/${databaseId}/metadata`));
    const names = new Set((meta.tables || []).map(t => (t.name || '').toLowerCase()));
    if (names.has('messageworkspace')) {
      type = 'messaging';                        // Messaging takes priority
    } else if (names.has('emailworkspace')) {
      type = 'email';
    } else if (names.has('filefolderinfo') || names.has('moveworkspaces')) {
      type = 'content';
    }
  } catch (err) {
    console.error(`Type detection failed for db ${databaseId}:`, err.message);
  }
  dbTypeCache.set(databaseId, type);
  return type;
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

// MongoDB errors that mean the database server is briefly unavailable (failover,
// a replica set member recovering, network blips). These are worth retrying.
const TRANSIENT_MONGO_ERROR = /NotPrimaryOrSecondary|node is recovering|NotWritablePrimary|NotMaster|PrimarySteppedDown|InterruptedDueToReplStateChange|HostUnreachable|SocketException|connection .*(closed|reset)/i;

// MongoDB errors that mean the server cannot be reached at all — it is either
// decommissioned, not started yet, or the port is blocked.
const SERVER_UNREACHABLE_ERROR = /Timed out while waiting for a server|MongoSocketOpenException|Exception opening socket|ServerSelectionTimeout|connection timed out|fetch failed/i;

const RETRY_DELAYS_MS = [1500, 4000];

// Maximum time to wait for a single collection query before giving up
const QUERY_TIMEOUT_MS = 50000;
// Email file collections (emailInfo, AttachMentDetails) can be very large — allow more time
const EMAIL_FILE_TIMEOUT_MS = 180000;

const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));

// Rejects after ms milliseconds with a timeout error
function withTimeout(promise, ms) {
  return new Promise((resolve, reject) => {
    const t = setTimeout(() => reject(new Error(`Query timed out after ${ms / 1000}s`)), ms);
    promise.then(v => { clearTimeout(t); resolve(v); }, e => { clearTimeout(t); reject(e); });
  });
}

// Turns raw Metabase/MongoDB errors into a short message for the dashboard
function friendlyError(err) {
  const msg = String(err?.message || err);
  if (/Query timed out after/.test(msg)) {
    return 'Query timed out — the collection may be very large. Try Refresh or check the migration server.';
  }
  if (SERVER_UNREACHABLE_ERROR.test(msg)) {
    return 'Server may be decommissioned or not started yet. Verify the migration server is running and try Refresh.';
  }
  if (TRANSIENT_MONGO_ERROR.test(msg)) {
    return 'The MongoDB server for this database is temporarily unavailable. This is a database server issue, not the dashboard. Try Refresh in a few minutes.';
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

// IDs (as strings) of the users whose email is not in an excluded domain.
// emailField defaults to 'primaryEmail' (Messaging/Content); pass 'email' for Email databases.
async function findIncludedUserIds(databaseId, emailField = 'primaryEmail') {
  const rows = await runMongoQuery(databaseId, 'Users', [
    { $match: { [emailField]: { $not: EXCLUDED_OWNER_REGEX } } },
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

// ---- Session ----
app.use(session({
  secret: SESSION_SECRET,
  resave: false,
  saveUninitialized: false,
  cookie: { httpOnly: true, secure: false, maxAge: 8 * 60 * 60 * 1000 } // 8 hours
}));

// ---- Auth routes (no auth required on these) ----
app.get('/auth/login', async (req, res) => {
  if (!msalApp) return res.redirect('/');
  try {
    const url = await msalApp.getAuthCodeUrl({
      scopes: ['openid', 'profile', 'email'],
      redirectUri: AZURE_REDIRECT_URI,
    });
    res.redirect(url);
  } catch (err) {
    res.redirect(`/login.html?error=${encodeURIComponent(err.message)}`);
  }
});

app.get('/auth/callback', async (req, res) => {
  const { code, error, error_description } = req.query;
  if (error) {
    console.error('Azure OAuth error:', error, error_description);
    const msg = error_description || error;
    return res.redirect(`/login.html?error=${encodeURIComponent(msg)}`);
  }
  try {
    const result = await msalApp.acquireTokenByCode({
      code,
      scopes: ['openid', 'profile', 'email'],
      redirectUri: AZURE_REDIRECT_URI,
    });
    const email = (result.account?.username || '').toLowerCase();
    if (!email.endsWith(`@${ALLOWED_DOMAIN}`)) {
      return res.redirect(`/login.html?error=${encodeURIComponent(`Access restricted to @${ALLOWED_DOMAIN} accounts.`)}`);
    }
    req.session.user = { name: result.account?.name || email, email };
    res.redirect('/');
  } catch (err) {
    res.redirect(`/login.html?error=${encodeURIComponent(err.message)}`);
  }
});

app.get('/auth/logout', (req, res) => {
  req.session.destroy(() => {
    const post = encodeURIComponent(AZURE_REDIRECT_URI.replace('/auth/callback', '/login.html'));
    res.redirect(`https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/logout?post_logout_redirect_uri=${post}`);
  });
});

app.get('/api/me', (req, res) => {
  if (!req.session?.user) return res.status(401).json({ error: 'Not authenticated' });
  res.json(req.session.user);
});

// ---- Auth guard ----
app.use((req, res, next) => {
  if (!msalApp) return next();               // SSO not configured — open access
  if (req.session?.user) return next();
  if (req.path === '/login.html' || req.path.startsWith('/auth/')) return next();
  if (req.path.startsWith('/api/')) return res.status(401).json({ error: 'Not authenticated' });
  res.redirect('/login.html');
});

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public'), {
  setHeaders: (res, filePath) => {
    if (filePath.endsWith('.html')) {
      res.setHeader('Cache-Control', 'no-store');
    }
  }
}));

async function loadMessagingDatabaseStatus(name, databaseId) {
  const workspaceBase = EXCLUDED_OWNER_REGEX ? { ownerEmailId: { $not: EXCLUDED_OWNER_REGEX } } : null;
  const userIdsPromise = EXCLUDED_OWNER_REGEX
    ? findIncludedUserIds(databaseId).catch(err => { console.error(`User lookup failed for ${name}:`, err.message); return null; })
    : Promise.resolve(null);

  async function queryWS(label, dmFilter) {
    try {
      const match = dmFilter !== undefined
        ? (workspaceBase ? { ...workspaceBase, directOrGroupMessage: dmFilter } : { directOrGroupMessage: dmFilter })
        : workspaceBase;
      return { collection: label, rows: await withTimeout(runProcessStatusQuery(databaseId, 'MessageWorkSpace', match), QUERY_TIMEOUT_MS) };
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
      return { collection: label, rows: await withTimeout(runProcessStatusQuery(databaseId, 'MessageEachFiles', match), QUERY_TIMEOUT_MS) };
    } catch (err) {
      console.error(`Query failed for ${name}.MessageEachFiles (${label}):`, err.message);
      return { collection: label, error: friendlyError(err) };
    }
  }

  const [wsTotal, wsChannels, wsDMs, efTotal, efChannels, efDMs] = await Promise.all([
    queryWS('MessageWorkSpace', undefined),
    queryWS('Channels',         false),
    queryWS('Direct Messages',  true),
    queryEF('MessageEachFiles', undefined),
    queryEF('Channels',         false),
    queryEF('Direct Messages',  true),
  ]);

  const allResults = [wsTotal, wsChannels, wsDMs, efTotal, efChannels, efDMs];
  const errors = allResults.map(c => c.error);
  if (errors.every(e => e) && new Set(errors).size === 1) {
    return { database: name, error: errors[0] };
  }

  let combinations = [];
  try {
    combinations = await findCombinations(databaseId, 'MessageWorkSpace', workspaceBase, 'fromCloudName', 'toCloudName');
  } catch (err) {
    console.error(`Combination lookup failed for ${name}:`, err.message);
  }

  return {
    database: name,
    projectType: 'messaging',
    combinations,
    workspace: [wsTotal, wsChannels, wsDMs],
    files:     [efTotal, efChannels, efDMs],
  };
}

async function loadContentDatabaseStatus(name, databaseId) {
  const workspaceBase = EXCLUDED_OWNER_REGEX ? { ownerEmailId: { $not: EXCLUDED_OWNER_REGEX } } : null;
  const userIdsPromise = EXCLUDED_OWNER_REGEX
    ? findIncludedUserIds(databaseId).catch(err => { console.error(`User lookup failed for ${name}:`, err.message); return null; })
    : Promise.resolve(null);

  async function querySimple(label, collection, match) {
    try {
      return { collection: label, rows: await withTimeout(runProcessStatusQuery(databaseId, collection, match), QUERY_TIMEOUT_MS) };
    } catch (err) {
      console.error(`Query failed for ${name}.${collection}:`, err.message);
      return { collection: label, error: friendlyError(err) };
    }
  }

  async function queryByUserId(label, collection) {
    try {
      const ids = await userIdsPromise;
      const match = ids ? { userId: { $in: ids } } : null;
      return { collection: label, rows: await withTimeout(runProcessStatusQuery(databaseId, collection, match), QUERY_TIMEOUT_MS) };
    } catch (err) {
      console.error(`Query failed for ${name}.${collection}:`, err.message);
      return { collection: label, error: friendlyError(err) };
    }
  }

  async function queryDriveChanges() {
    try {
      // Exclude _id only; return all fields so the frontend can pick them
      // regardless of exact casing in this database's collection
      const pipeline = [
        { $project: { _id: 0 } },
        { $limit: 200 }
      ];
      const rows = await withTimeout(runMongoQuery(databaseId, 'DriveChangeIdDetails', pipeline), QUERY_TIMEOUT_MS);
      return { collection: 'DriveChangeIdDetails', rows };
    } catch (err) {
      console.error(`Query failed for ${name}.DriveChangeIdDetails:`, err.message);
      return { collection: 'DriveChangeIdDetails', error: friendlyError(err) };
    }
  }

  const [moveWS, fileFolderInfo, folderMetadataInfo, collaborationDetails, hyperlinks, driveChanges] = await Promise.all([
    querySimple('MoveWorkSpaces',        'MoveWorkSpaces',        workspaceBase),
    queryByUserId('FileFolderInfo',       'FileFolderInfo'),
    queryByUserId('FolderMetadataInfo',   'FolderMetadataInfo'),
    querySimple('CollabarationDetails',  'CollabarationDetails',  null),
    querySimple('HyperLinks',            'HyperLinks',            null),
    queryDriveChanges(),
  ]);

  const allResults = [moveWS, fileFolderInfo, folderMetadataInfo, collaborationDetails, hyperlinks];
  const errors = allResults.map(c => c.error);
  if (errors.every(e => e) && new Set(errors).size === 1) {
    return { database: name, error: errors[0] };
  }

  let combinations = [];
  try {
    combinations = await findCombinations(databaseId, 'MoveWorkSpaces', workspaceBase, 'fromCloudName', 'toCloudName');
  } catch (err) {
    console.error(`Combination lookup failed for ${name}:`, err.message);
  }

  return {
    database: name,
    projectType: 'content',
    combinations,
    workspace: [moveWS],
    files:     [fileFolderInfo, folderMetadataInfo, collaborationDetails, hyperlinks],
    driveChanges,
  };
}

async function loadEmailDatabaseStatus(name, databaseId) {
  const workspaceBase = EXCLUDED_OWNER_REGEX ? { ownerEmailId: { $not: EXCLUDED_OWNER_REGEX } } : null;

  // Collect the _id values of non-internal workspaces so we can filter
  // emailInfo and AttachMentDetails by emailWorkSpaceId (much smaller than userId $in)
  const wsIdsPromise = EXCLUDED_OWNER_REGEX
    ? runMongoQuery(databaseId, 'emailWorkSpace', [
        ...(workspaceBase ? [{ $match: workspaceBase }] : []),
        { $group: { _id: null, ids: { $push: { $toString: '$_id' } } } }
      ]).then(rows => {
        const ids = rows[0]?.ids || [];
        return typeof ids === 'string' ? JSON.parse(ids) : ids;
      }).catch(err => { console.error(`Workspace ID lookup failed for ${name}:`, err.message); return null; })
    : Promise.resolve(null);

  async function queryCollection(label, collection, match, timeoutMs = QUERY_TIMEOUT_MS) {
    try {
      return { collection: label, rows: await withTimeout(runProcessStatusQuery(databaseId, collection, match), timeoutMs) };
    } catch (err) {
      console.error(`Query failed for ${name}.${collection}:`, err.message);
      return { collection: label, error: friendlyError(err) };
    }
  }

  async function queryByWorkspaceId(label, collection) {
    try {
      const wsIds = await wsIdsPromise;
      const match = wsIds?.length ? { emailWorkSpaceId: { $in: wsIds } } : null;
      return { collection: label, rows: await withTimeout(runProcessStatusQuery(databaseId, collection, match), EMAIL_FILE_TIMEOUT_MS) };
    } catch (err) {
      console.error(`Query failed for ${name}.${collection}:`, err.message);
      return { collection: label, error: friendlyError(err) };
    }
  }

  // Run emailWorkSpace first (fast), then emailInfo and AttachMentDetails sequentially.
  // AttachMentDetails uses no filter — a plain $group scan is faster than a
  // non-indexed $in match on a very large collection.
  const emailWS = await queryCollection('Email Workspace', 'emailWorkSpace', workspaceBase);
  const emailInfo = await queryByWorkspaceId('Email Info', 'emailInfo');
  const attachmentDetails = await queryCollection('Attachment Details', 'AttachMentDetails', null, EMAIL_FILE_TIMEOUT_MS);

  const allResults = [emailWS, emailInfo, attachmentDetails];
  const errors = allResults.map(c => c.error);
  if (errors.every(e => e) && new Set(errors).size === 1) {
    return { database: name, error: errors[0] };
  }

  let combinations = [];
  try {
    combinations = await findCombinations(databaseId, 'emailWorkSpace', workspaceBase, 'fromCloud', 'toCloud');
  } catch (err) {
    console.error(`Combination lookup failed for ${name}:`, err.message);
  }

  return {
    database: name,
    projectType: 'email',
    combinations,
    workspace: [emailWS],
    files:     [emailInfo, attachmentDetails],
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
  const type = await detectDatabaseType(databaseId);
  let result;
  if (type === 'content') result = await loadContentDatabaseStatus(name, databaseId);
  else if (type === 'email') result = await loadEmailDatabaseStatus(name, databaseId);
  else result = await loadMessagingDatabaseStatus(name, databaseId);
  return applySnapshots(name, result);
}

// Source -> destination cloud pairs from MessageWorkSpace, most workspaces first.
// Uses the same owner filter as the counts; falls back to all workspaces when
// the filter leaves none (e.g. an internal-only test database).
async function findCombinations(databaseId, collection, workspaceMatch, fromField = 'fromCloudName', toField = 'toCloudName') {
  const pipeline = match => [
    ...(match ? [{ $match: match }] : []),
    { $group: { _id: { from: `$${fromField}`, to: `$${toField}` }, Count: { $sum: 1 } } },
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
    if (req.body?.interval !== undefined) {
      const mins = Number(req.body.interval);
      if (![0, 30, 60, 120, 180].includes(mins)) throw new Error('interval must be 0, 30, 60, 120, or 180 minutes');
      if (mins > 0 && !next.webhookUrl) throw new Error('Add the Teams webhook URL before enabling auto-posts');
      next.interval = mins;
      delete next.hourly; // migrate away from legacy field
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

// Fires every 30 minutes (at :00 and :30). For each project checks if the current half-hour
// slot aligns with the configured interval (30, 60, 120, or 180 minutes).
async function sendScheduledUpdates() {
  const now = new Date();
  const minuteOfDay = now.getHours() * 60 + now.getMinutes(); // 0‥1439
  const due = Object.keys(projectSettings).filter(p => {
    const s = projectSettings[p];
    const mins = s.interval ?? (s.hourly ? 60 : 0);
    if (!mins || !s.webhookUrl) return false;
    return minuteOfDay % mins === 0;
  });
  for (const project of due) {
    try {
      await sendProjectStatus(project);
    } catch (err) {
      console.error(`Scheduled Teams update failed for ${project}:`, err.message);
    }
  }
}

function scheduleUpdates() {
  const now = new Date();
  const next = new Date(now);
  // Advance to the next :00 or :30 mark
  const m = now.getMinutes();
  const addMins = m < 30 ? (30 - m) : (60 - m);
  next.setMinutes(now.getMinutes() + addMins, 0, 0);
  setTimeout(async () => {
    await sendScheduledUpdates();
    scheduleUpdates();
  }, next - now);
  console.log(`Next scheduled Teams check at ${next.toLocaleTimeString()}`);
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
  scheduleUpdates();
});
