// Posting project status summaries to Microsoft Teams group chats.
//
// Each project has its own Teams "Workflows" webhook URL (created in the group
// chat with the "Send webhook alerts to a chat" workflow). We POST an Adaptive
// Card to it and the Workflows bot posts the card into the chat.

const fmt = new Intl.NumberFormat('en-US');

const WS_COLLECTIONS = ['MessageWorkSpace', 'Channels', 'Direct Messages'];
const EF_COLLECTIONS = ['MessageEachFiles', 'Channels', 'Direct Messages'];

// Only Microsoft webhook hosts are accepted, so a saved URL can't be used to
// make the server call arbitrary addresses
const ALLOWED_WEBHOOK_HOST = /(^|\.)(logic\.azure\.com|powerplatform\.com|powerautomate\.com|webhook\.office\.com)$/i;

// Links to a workflow's page rather than its webhook: a common wrong paste
const WORKFLOW_PAGE_HOST = /(^|\.)(teams\.microsoft\.com|teams\.cloud\.microsoft|make\.powerautomate\.com|flow\.microsoft\.com|powerautomate\.microsoft\.com)$/i;

function validateWebhookUrl(value) {
  let url;
  try {
    url = new URL(String(value || '').trim().replace(/^["'<]+|["'>]+$/g, ''));
  } catch {
    throw new Error('That is not a valid URL. Copy the whole webhook URL, starting with https://');
  }
  if (url.protocol !== 'https:') throw new Error('The webhook URL must start with https://');
  const host = url.hostname.replace(/\.$/, '');
  if (WORKFLOW_PAGE_HOST.test(host)) {
    throw new Error(`This is a link to the workflow page (${host}), not its webhook URL. ` +
      'Open the workflow, click its first step “When a Teams webhook request is received”, and copy the HTTP URL shown there.');
  }
  if (!ALLOWED_WEBHOOK_HOST.test(host)) {
    console.warn(`Rejected Teams webhook host: ${host}`);
    throw new Error(`This URL points to ${host}, which is not a Teams Workflows webhook. ` +
      'The webhook URL ends in logic.azure.com or powerplatform.com; copy it from the workflow’s first step.');
  }
  return url.toString();
}

// Shows enough of the URL to recognise it without exposing the signature
function maskWebhookUrl(value) {
  if (!value) return '';
  try {
    const url = new URL(value);
    return `${url.protocol}//${url.hostname}/…${value.slice(-6)}`;
  } catch {
    return '…';
  }
}

// Only non-default properties are kept, to stay well under Teams' card size limit
function text(value, opts = {}) {
  const block = { type: 'TextBlock', text: String(value), wrap: true, spacing: 'None', ...opts };
  for (const key of Object.keys(block)) {
    if (block[key] === false || block[key] === 'Default' || block[key] === 'Left') delete block[key];
  }
  return block;
}

const CLOUD_LABELS = { MICROSOFT_TEAMS: 'Teams', SLACK: 'Slack', GOOGLE_CHAT: 'Google Chat' };
function cloudLabel(name) {
  if (!name) return 'Unknown';
  return CLOUD_LABELS[name] ||
    String(name).toLowerCase().split('_').map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ');
}
const comboLabel = c => `${cloudLabel(c.from)} to ${cloudLabel(c.to)}`;

const isDone = s => /^PROCESSED/.test(String(s ?? ''));
const isConflict = s => /CONFLICT/.test(String(s ?? '')) && !isDone(s);

function stats(rows) {
  let total = 0, done = 0, conflicts = 0;
  for (const [s, c] of rows) {
    const n = Number(c) || 0;
    total += n;
    if (isDone(s)) done += n;
    if (isConflict(s)) conflicts += n;
  }
  return { total, done, conflicts };
}

// A row of small cells; ColumnSets render in every Teams client, unlike the Table element
function row(cells, widths, { header = false, bold = false, separator = false } = {}) {
  const set = {
    type: 'ColumnSet',
    spacing: 'Small',
    columns: cells.map((cell, i) => ({
      type: 'Column',
      width: widths[i],
      items: [text(cell, {
        size: 'Small',
        weight: header || bold ? 'Bolder' : 'Default',
        color: header ? 'Accent' : 'Default',
        horizontalAlignment: i === 0 ? 'Left' : 'Right'
      })]
    }))
  };
  if (separator) set.separator = true;
  return set;
}

const STATUS_WIDTHS = [3, 2];

// One collection's statuses for one database: ProcessStatus | Count, largest first
function collectionColumn(c) {
  const items = [text(c.collection.toUpperCase(), { size: 'Small', weight: 'Bolder', isSubtle: true })];
  if (c.error) {
    items.push(text(c.error, { size: 'Small', color: 'Attention', spacing: 'Small' }));
    return { type: 'Column', width: 1, items };
  }
  const rows = [...c.rows].sort((a, b) => (Number(b[1]) || 0) - (Number(a[1]) || 0));
  const { total } = stats(rows);
  items.push(row(['ProcessStatus', 'Count'], STATUS_WIDTHS, { header: true }));
  if (!rows.length) items.push(text('No records', { size: 'Small', isSubtle: true, spacing: 'Small' }));
  for (const [status, count] of rows) {
    const n = Number(count) || 0;
    items.push(row([status ?? '(null)', fmt.format(n)], STATUS_WIDTHS));
  }
  items.push(row(['Total', fmt.format(total)], STATUS_WIDTHS, { bold: true, separator: true }));
  return { type: 'Column', width: 1, items };
}

// One row of collection columns (workspace or files)
function collectionRow(arr, names) {
  const cols = names
    .map((name, i) => arr[i] || null)
    .filter(Boolean)
    .map(collectionColumn)
    .map((col, i) => (i ? { ...col, spacing: 'ExtraLarge', separator: true } : col));
  return { type: 'ColumnSet', spacing: 'Medium', columns: cols };
}

// One database per block: workspace row (WS | Channels | DMs) then files row
function databaseBlock(r) {
  const items = [];
  const combos = (r.combinations || []).map(comboLabel);
  items.push(text([r.database, ...combos].join(' · '), { weight: 'Bolder', size: 'Medium' }));
  if (r.error) {
    items.push(text(r.error, { size: 'Small', color: 'Attention', spacing: 'Small' }));
  } else {
    items.push(collectionRow(r.workspace || [], WS_COLLECTIONS));
    items.push({ ...collectionRow(r.files || [], EF_COLLECTIONS), separator: true, spacing: 'Medium' });
  }
  return { type: 'Container', separator: true, spacing: 'Large', items };
}

// Project totals across its databases: database count, workspace and files totals
function projectSummary(results) {
  const summaryDefs = [
    { name: 'MessageWorkSpace', get: r => r.workspace && r.workspace[0] },
    { name: 'MessageEachFiles', get: r => r.files && r.files[0] },
  ];
  const merged = summaryDefs.map(({ name, get }) => {
    const m = new Map();
    for (const r of results) {
      const c = get(r);
      if (c && !c.error) for (const [s, n] of c.rows) m.set(s, (m.get(s) || 0) + (Number(n) || 0));
    }
    return { name, ...stats([...m.entries()]) };
  });
  const loaded = results.filter(r => !r.error).length;
  const fact = (label, value, detail) => ({
    type: 'Column',
    width: 1,
    items: [
      text(label.toUpperCase(), { size: 'Small', isSubtle: true, weight: 'Bolder' }),
      text(value, { size: 'Large', weight: 'Bolder', spacing: 'Small' }),
      text(detail, { size: 'Small', isSubtle: true })
    ]
  });
  return {
    type: 'ColumnSet',
    spacing: 'Medium',
    columns: [
      fact('Databases', String(results.length),
        loaded === results.length ? 'All loaded' : `${loaded} loaded · ${results.length - loaded} failed`),
      ...merged.map(m => fact(m.name, fmt.format(m.total),
        `Processed ${fmt.format(m.done)} · Conflicts ${fmt.format(m.conflicts)}`))
    ]
  };
}

function buildCard(project, results, { excludedDomains = [] } = {}) {
  const now = new Date();
  const combinations = [...new Set(results.flatMap(r => (r.combinations || []).map(comboLabel)))];
  const subtitle = [
    combinations.length ? combinations.join(', ') : null,
    `${results.length} database${results.length === 1 ? '' : 's'}`,
    now.toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' })
  ].filter(Boolean).join(' · ');

  const body = [
    text(`${project} – Migration status`, { size: 'Large', weight: 'Bolder' }),
    text(subtitle, { isSubtle: true, size: 'Small', spacing: 'Small' }),
    projectSummary(results),
    ...results.map(databaseBlock)
  ];

  if (excludedDomains.length) {
    body.push(text(`Excludes users and workspaces owned by ${excludedDomains.map(d => '@' + d).join(', ')}`, {
      isSubtle: true, size: 'Small', spacing: 'Medium'
    }));
  }

  return {
    type: 'message',
    attachments: [{
      contentType: 'application/vnd.microsoft.card.adaptive',
      contentUrl: null,
      content: {
        $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
        type: 'AdaptiveCard',
        version: '1.4',
        msteams: { width: 'Full' },
        body
      }
    }]
  };
}

async function postToWebhook(webhookUrl, payload) {
  const cardJson = JSON.stringify(payload);
  console.log(`Teams card size: ${cardJson.length} bytes`);
  const res = await fetch(webhookUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: cardJson
  });
  const responseBody = (await res.text()).slice(0, 500);
  console.log(`Teams webhook response: ${res.status} ${responseBody}`);
  if (!res.ok) {
    throw new Error(`Teams webhook returned ${res.status}${responseBody ? ': ' + responseBody : ''}`);
  }
}

module.exports = { validateWebhookUrl, maskWebhookUrl, buildCard, postToWebhook };
