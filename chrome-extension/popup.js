const statusEl = document.getElementById('status');
const sizeEl = document.getElementById('leagueSize');

async function activeDraftTab() {
  const tabs = await chrome.tabs.query({active: true, currentWindow: true});
  return tabs.find(t => /^https:\/\/fantasy\.espn\.com\/football\/draft/i.test(t.url || '')) || null;
}

async function getStats() {
  return await chrome.runtime.sendMessage({type:'get-stats'}).catch(() => ({}));
}

async function check() {
  statusEl.textContent = 'Checking bridge + Excel…';
  const result = await chrome.runtime.sendMessage({type:'bridge-status'}).catch(e => ({ok:false,error:String(e)}));
  const s = result?.stats || await getStats();
  const tab = await activeDraftTab();
  let armed = false;
  if (tab?.id) {
    const arm = await chrome.tabs.sendMessage(tab.id, {type:'ticker-ready-state'}).catch(() => null);
    armed = !!arm?.armed;
  }

  if (result?.ok && armed) {
    statusEl.innerHTML = `<div class="readyBig">READY FOR DRAFT ✓</div>` +
      `Bridge: connected &nbsp; Excel: connected<br>` +
      `Ticker: ARMED + RECONCILING &nbsp; League size: ${s.leagueSize || '?'}<br>` +
      `${s.lastPick ? `Last: ${s.lastPick}` : 'Waiting for pick #1…'}`;
  } else if (result?.ok) {
    statusEl.innerHTML = `<span class="warn">CONNECTING…</span><br>` +
      `Bridge + Excel are reachable. Waiting for the ESPN ticker listener to arm.<br>` +
      `Do not start the draft until this says READY FOR DRAFT.`;
  } else {
    statusEl.innerHTML = `<span class="bad">NOT READY</span><br>` +
      `${result?.error || 'Start the local bridge and open Excel.'}<br>` +
      `<b>Do not start the draft yet.</b>`;
  }
}

document.getElementById('save').addEventListener('click', async () => {
  const n = Number(sizeEl.value);
  if (!Number.isInteger(n) || n < 4 || n > 20) {
    statusEl.innerHTML = '<span class="bad">Enter a league size from 4 to 20.</span>';
    return;
  }

  await chrome.storage.local.set({leagueSize:n});
  const tab = await activeDraftTab();
  if (tab?.id) {
    await chrome.tabs.sendMessage(tab.id, {type:'ticker-rescan'}).catch(() => {});
  }
  setTimeout(check, 150);
});

document.getElementById('check').addEventListener('click', check);

chrome.storage.local.get({leagueSize:10}).then(v => {
  sizeEl.value = v.leagueSize || 10;
  check();
  setInterval(check, 750);
});
