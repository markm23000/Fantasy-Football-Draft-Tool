const BASE = 'http://127.0.0.1:8765';
const STORE_KEY = 'espnExcelTickerPendingV24';
const LEGACY_KEYS = ['espnExcelTickerPendingV22'];
let stats = {version: '2.4-ticker-reconcile-fixed', leagueSize: 10, parsedCards: 0, sentPicks: 0, lastPick: '', lastScanAt: 0, ready: false, readyDetail: 'Checking bridge + Excel…'};
let pending = [];
let posting = false;
let hydrated = false;
let hydratePromise = null;

function pickNumber(p) {
  const n = Number(p?.pick ?? p?.overall);
  return Number.isInteger(n) && n > 0 ? n : null;
}

function idOf(p) {
  const n = pickNumber(p);
  if (n !== null) return `pick:${n}`;
  return `fallback:${p?.round || ''}:${p?.round_pick || ''}:${String(p?.player || '').toLowerCase()}`;
}

function signatureOf(p) {
  return `${idOf(p)}|${String(p?.player || '').trim().toLowerCase()}|${String(p?.drafter || '').trim().toLowerCase()}`;
}

// One pending record per numbered pick. If ESPN later supplies a correction or
// team name for the same pick, the newest version replaces the older one.
function normalizePending(items) {
  const byId = new Map();
  for (const p of items || []) {
    if (!p || !p.player) continue;
    byId.set(idOf(p), p);
  }
  return [...byId.values()].sort((a, b) => {
    const ap = pickNumber(a);
    const bp = pickNumber(b);
    if (ap !== null && bp !== null) return ap - bp;
    if (ap !== null) return -1;
    if (bp !== null) return 1;
    return idOf(a).localeCompare(idOf(b));
  });
}

async function hydrate() {
  if (hydrated) return;
  if (hydratePromise) return hydratePromise;
  hydratePromise = (async () => {
    const defaults = {[STORE_KEY]: []};
    for (const k of LEGACY_KEYS) defaults[k] = [];
    const saved = await chrome.storage.local.get(defaults);
    const legacy = LEGACY_KEYS.flatMap(k => saved[k] || []);
    pending = normalizePending([...legacy, ...(saved[STORE_KEY] || []), ...pending]);
    hydrated = true;
    await savePending();
    if (LEGACY_KEYS.length) await chrome.storage.local.remove(LEGACY_KEYS).catch(() => {});
  })();
  try { await hydratePromise; } finally { hydratePromise = null; }
}

async function savePending() {
  await chrome.storage.local.set({[STORE_KEY]: pending});
}

async function postJSON(path, body) {
  const resp = await fetch(BASE + path, {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(body || {})
  });
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok) throw new Error(data.error || `HTTP ${resp.status}`);
  return data;
}

async function enqueue(picks) {
  await hydrate();
  pending = normalizePending([...pending, ...(picks || [])]);
  // Persist before acknowledging the content script so early picks survive an
  // MV3 service-worker sleep, Chrome restart, or momentary Excel disconnect.
  await savePending();
  flush();
}

async function flush() {
  await hydrate();
  if (posting || !pending.length) return;
  posting = true;
  const batch = normalizePending(pending.slice());
  const acknowledged = new Map(batch.map(p => [idOf(p), signatureOf(p)]));
  try {
    await postJSON('/picks', {picks: batch});
    // Remove only the exact version Excel acknowledged. If a correction for the
    // same pick arrived while this POST was in flight, keep that newer version.
    pending = pending.filter(p => acknowledged.get(idOf(p)) !== signatureOf(p));
    pending = normalizePending(pending);
    await savePending();
  } catch (_) {
    // Keep durable queue untouched. Heartbeats/new picks/popup checks retry it.
  } finally {
    posting = false;
  }
  if (pending.length) setTimeout(flush, 200);
}

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg?.type === 'ticker-picks' && Array.isArray(msg.picks)) {
    enqueue(msg.picks)
      .then(() => sendResponse({ok: true, queued: msg.picks.length}))
      .catch(e => sendResponse({ok: false, error: String(e?.message || e)}));
    return true;
  }
  if (msg?.type === 'ticker-heartbeat') {
    flush();
    sendResponse({ok: true, pending: pending.length});
    return true;
  }
  if (msg?.type === 'ticker-stats') {
    stats = {...stats, ...(msg.stats || {})};
    postJSON('/stats', {...stats, pending: pending.length}).catch(() => {});
    flush();
    sendResponse({ok: true});
    return true;
  }
  if (msg?.type === 'get-stats') {
    hydrate().then(() => sendResponse({...stats, pending: pending.length}));
    return true;
  }
  if (msg?.type === 'bridge-status') {
    flush();
    fetch(BASE + '/status').then(async r => {
      const data = await r.json().catch(() => ({}));
      if (!r.ok) throw new Error(data.error || `HTTP ${r.status}`);
      await hydrate();
      sendResponse({ok: true, data, stats: {...stats, pending: pending.length}});
    }).catch(async e => {
      await hydrate().catch(() => {});
      sendResponse({ok: false, error: String(e?.message || e), stats: {...stats, pending: pending.length}});
    });
    return true;
  }
  if (msg?.type === 'bridge-refresh') {
    postJSON('/refresh', {}).then(data => {
      flush();
      sendResponse({ok: true, data});
    }).catch(e => sendResponse({ok: false, error: String(e?.message || e)}));
    return true;
  }
});

hydrate().then(flush).catch(() => {});
