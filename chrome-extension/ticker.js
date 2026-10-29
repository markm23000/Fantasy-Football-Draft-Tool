(() => {
  'use strict';
  if (window.__espnExcelTickerOnlyV24) return;
  window.__espnExcelTickerOnlyV24 = true;

  const VERSION = '2.4-ticker-reconcile-fixed';
  const seenByPick = new Map();
  const inflightByPick = new Map();
  let leagueSize = 10;
  let parsedCards = 0;
  let sentPicks = 0;
  let lastPick = '';
  let lastScanAt = 0;
  let scanTimer = null;
  let followupTimer = null;

  const PICK_RE = /R\s*(\d{1,2})\s*,\s*P\s*(\d{1,2})\s*-\s*([^\n\r]+)/i;
  const PICK_RE_GLOBAL = /R\s*\d{1,2}\s*,\s*P\s*\d{1,2}\s*-/ig;
  const PLAYER_LINE_RE = /^(.+?)\s*\/\s*[A-Z]{2,4}\s+(QB|RB|WR|TE|K|DST|D\/ST)\b/i;

  function cleanText(v) {
    return String(v || '').replace(/\u00a0/g, ' ').replace(/[ \t]+/g, ' ').trim();
  }

  function visible(el) {
    if (!(el instanceof Element) || !el.isConnected) return false;
    const r = el.getBoundingClientRect();
    if (r.width <= 0 || r.height <= 0) return false;
    const st = getComputedStyle(el);
    return st.display !== 'none' && st.visibility !== 'hidden';
  }

  function elementText(el) {
    // innerText is best for rendered cards; textContent rescues cards during
    // React/ESPN transitions before they become visible or just after removal.
    return String(el?.innerText || el?.textContent || '');
  }

  function extractFromText(raw) {
    const text = String(raw || '').replace(/\r/g, '');
    const marker = text.match(PICK_RE);
    if (!marker) return null;

    const round = Number(marker[1]);
    const roundPick = Number(marker[2]);
    if (!Number.isInteger(round) || !Number.isInteger(roundPick) || round < 1 || round > 30) return null;
    if (!Number.isInteger(leagueSize) || leagueSize < 4 || leagueSize > 20) return null;
    if (roundPick < 1 || roundPick > leagueSize) return null;

    const prefix = text.slice(0, marker.index);
    const lines = prefix.split('\n').map(cleanText).filter(Boolean);
    let player = '';

    for (let i = lines.length - 1; i >= 0; i--) {
      const m = lines[i].match(PLAYER_LINE_RE);
      if (m) {
        player = cleanText(m[1]);
        break;
      }
    }

    if (!player) return null;
    if (/^picks?$/i.test(player) || player.length < 2 || player.length > 80) return null;

    const overall = (round - 1) * leagueSize + roundPick;
    const drafter = cleanText(marker[3]);
    return {pick: overall, player, round, round_pick: roundPick, drafter, source: 'picks-ticker'};
  }

  function parseCandidate(el) {
    if (!(el instanceof Element)) return null;
    const text = elementText(el);
    if (text.length < 12 || text.length > 320) return null;
    const markers = text.match(PICK_RE_GLOBAL);
    if (!markers || markers.length !== 1) return null;
    if (!text.includes('/')) return null;
    return extractFromText(text);
  }

  function chooseCards(root) {
    const found = new Map();
    const nodes = [];
    if (root instanceof Element) nodes.push(root);
    if (root && root.querySelectorAll) {
      for (const el of root.querySelectorAll('div, li, article')) nodes.push(el);
    }

    for (const el of nodes) {
      const p = parseCandidate(el);
      if (!p) continue;
      // Pick number is the identity. Prefer a visible card, but never require
      // visibility; hidden/detached transition nodes can contain the pick we need.
      const score = (visible(el) ? 0 : 1000) + elementText(el).length;
      const prev = found.get(p.pick);
      if (!prev || score < prev.score) found.set(p.pick, {pick: p, score});
    }
    return [...found.values()].map(x => x.pick).sort((a, b) => a.pick - b.pick);
  }

  function signature(p) {
    return `${cleanText(p?.player).toLowerCase()}|${cleanText(p?.drafter).toLowerCase()}`;
  }

  function sendNew(picks) {
    const fresh = [];
    const meta = [];
    for (const p of picks) {
      const id = Number(p.pick);
      if (!Number.isInteger(id) || id <= 0) continue;
      const sig = signature(p);
      if (seenByPick.get(id) === sig || inflightByPick.get(id) === sig) continue;
      inflightByPick.set(id, sig);
      fresh.push(p);
      meta.push({id, sig});
    }
    if (!fresh.length) return;

    chrome.runtime.sendMessage({type: 'ticker-picks', picks: fresh}).then(resp => {
      if (!resp?.ok) return;
      for (let i = 0; i < fresh.length; i++) {
        const {id, sig} = meta[i];
        seenByPick.set(id, sig);
        parsedCards++;
        sentPicks++;
        lastPick = `#${fresh[i].pick} ${fresh[i].player}`;
      }
      publishStats();
    }).catch(() => {}).finally(() => {
      for (const {id, sig} of meta) {
        if (inflightByPick.get(id) === sig) inflightByPick.delete(id);
      }
    });
  }

  function scan(root = document.body) {
    if (!root) return;
    lastScanAt = Date.now();
    try { sendNew(chooseCards(root)); } catch (_) {}
  }

  function scheduleFullScan(delay = 35) {
    clearTimeout(scanTimer);
    scanTimer = setTimeout(() => scan(document.body), delay);

    // A second pass catches ESPN/React cards whose text lands shortly after the
    // DOM node itself. This is especially important for picks #1 and #2.
    clearTimeout(followupTimer);
    followupTimer = setTimeout(() => scan(document.body), Math.max(140, delay + 100));
  }

  function publishStats() {
    chrome.runtime.sendMessage({
      type: 'ticker-stats',
      stats: {
        version: VERSION,
        leagueSize,
        parsedCards,
        sentPicks,
        lastPick,
        lastScanAt
      }
    }).catch(() => {});
  }

  let armed = false;
  let observer = null;
  let readyTimer = null;

  function publishReady(ready, detail = '') {
    chrome.runtime.sendMessage({
      type: 'ticker-stats',
      stats: {
        version: VERSION,
        leagueSize,
        parsedCards,
        sentPicks,
        lastPick,
        lastScanAt,
        ready: !!ready,
        readyDetail: detail
      }
    }).catch(() => {});
  }

  function armTicker() {
    if (armed) return;
    armed = true;
    if (readyTimer) {
      clearInterval(readyTimer);
      readyTimer = null;
    }

    // Bridge + Excel are confirmed reachable. Arm before pick #1 and reconcile
    // the complete ticker several times during ESPN's initial render.
    scan(document.body);
    for (const delay of [75, 200, 450, 900, 1600, 3000]) {
      setTimeout(() => scan(document.body), delay);
    }

    observer = new MutationObserver((mutations) => {
      let changed = false;
      for (const m of mutations) {
        for (const node of m.addedNodes || []) {
          if (node instanceof Element) {
            // Parse the exact added subtree immediately, even before it is visible.
            scan(node);
            changed = true;
          }
        }
        if (m.type === 'characterData' && m.target?.parentElement) {
          scan(m.target.parentElement);
          changed = true;
        }
      }
      if (changed) scheduleFullScan(35);
    });

    if (document.body) {
      observer.observe(document.body, {childList: true, subtree: true, characterData: true});
    }

    // Lightweight backup reconciliation. The mutation observer handles speed;
    // this full scan handles anything ESPN changed in a way we did not observe.
    setInterval(() => {
      scan(document.body);
      chrome.runtime.sendMessage({type: 'ticker-heartbeat'}).catch(() => {});
    }, 1000);

    publishReady(true, 'Bridge + Excel confirmed. Ticker armed + reconciling.');
  }

  async function checkReady() {
    if (armed) return;
    try {
      const result = await chrome.runtime.sendMessage({type: 'bridge-status'});
      if (result?.ok && result?.data?.ok !== false) {
        armTicker();
      } else {
        publishReady(false, 'Waiting for local bridge + Excel…');
      }
    } catch (_) {
      publishReady(false, 'Waiting for local bridge + Excel…');
    }
  }

  chrome.storage.local.get({leagueSize: 10}).then(v => {
    const n = Number(v.leagueSize);
    if (Number.isInteger(n) && n >= 4 && n <= 20) leagueSize = n;
    publishReady(false, 'Checking bridge + Excel…');
    checkReady();
    readyTimer = setInterval(checkReady, 500);
  });

  chrome.storage.onChanged.addListener((changes, area) => {
    if (area !== 'local' || !changes.leagueSize) return;
    const n = Number(changes.leagueSize.newValue);
    if (Number.isInteger(n) && n >= 4 && n <= 20) {
      leagueSize = n;
      seenByPick.clear();
      inflightByPick.clear();
      if (armed) scheduleFullScan(0);
      publishReady(armed, armed ? 'Bridge + Excel confirmed. Ticker armed + reconciling.' : 'Waiting for local bridge + Excel…');
    }
  });

  chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
    if (msg?.type === 'ticker-rescan') {
      if (armed) scan(document.body);
      publishStats();
      sendResponse({ok: true, armed, leagueSize, parsedCards, sentPicks, lastPick});
      return true;
    }
    if (msg?.type === 'ticker-ready-state') {
      sendResponse({ok: true, armed, leagueSize, parsedCards, sentPicks, lastPick});
      return true;
    }
  });
})();
