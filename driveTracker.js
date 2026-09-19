// ==UserScript==
// @name         UMS Placement Drive Tracker
// @namespace    namit.ums.drive.tracker
// @version      4.0
// @description  Keeps your last 10 registered placement drives (from the UMS Drive Registration page) in an Excel file
// @match        https://ums.lpu.in/Placements/*
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_registerMenuCommand
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// ==/UserScript==

(function () {
  'use strict';

  /* ------------------------------ CONFIG ------------------------------ */
  const CFG = {
    maxDrives: 10,                       // how many registered drives to keep
    tableSelector: 'table[id$="gdvPlacement"]',
    registrationPageUrl: '/Placements/frmPlacementDriveRegistration.aspx',
    autoFetchEveryMin: 60,               // background refresh interval on other Placements pages
    fileName: 'placement_drives.xlsx',
    syncDebounceMs: 800,
  };

  const KEY = 'drives_v3';
  const DIRTY = 'dirty_v3';
  const LAST_FETCH = 'last_fetch_v3';
  const onRegistrationPage = /frmPlacementDriveRegistration/i.test(location.pathname);

  /* ------------------------------ helpers ------------------------------ */
  const norm = (s) => (s || '').replace(/\s+/g, ' ').trim();
  const today = () => new Date().toISOString().slice(0, 10);
  const load = () => { try { return JSON.parse(GM_getValue(KEY, '{}')); } catch (e) { return {}; } };
  const saveData = (d) => GM_setValue(KEY, JSON.stringify(d));

  // Drive Registration page columns (the home-page table has no "Drive Code", so it is ignored)
  const COLS = {
    code: /^drive code/i,
    company: /^company/i,
    registered: /^registered$/i,
  };

  function getCols(table) {
    const head = table.rows[0];
    if (!head) return null;
    const map = {};
    Array.from(head.cells).forEach((c, i) => {
      const t = norm(c.textContent);
      Object.keys(COLS).forEach((k) => {
        if (map[k] === undefined && COLS[k].test(t)) map[k] = i;
      });
    });
    return map.code !== undefined && map.company !== undefined && map.registered !== undefined ? map : null;
  }

  // Decide state from the text of the "Registered" cell
  //   "Click to Cancel Registration" -> registered
  //   "Click to Register"            -> not registered
  //   anything else                  -> unknown (stored drives are left untouched)
  function registrationState(text) {
    if (/cancel/i.test(text)) return 'registered';
    if (/click to register/i.test(text)) return 'not';
    if (/registered/i.test(text) && !/not registered/i.test(text)) return 'registered';
    return 'unknown';
  }

  function readEntries(doc) {
    const entries = [];
    doc.querySelectorAll(CFG.tableSelector).forEach((table) => {
      const map = getCols(table);
      if (!map) return;
      const maxIdx = Math.max.apply(null, Object.keys(map).map((k) => map[k]));
      Array.from(table.rows).slice(1).forEach((row) => {
        const c = row.cells;
        if (row.querySelector('th') || c.length <= maxIdx) return;
        const code = norm(c[map.code].textContent);
        if (!code) return;
        entries.push({
          code,
          company: norm(c[map.company].textContent),
          state: registrationState(norm(c[map.registered].textContent)),
        });
      });
    });
    return entries;
  }

  // Newest registration date first, then position on the portal page (top = newest)
  const cmp = (a, b) =>
    (b.registeredOn || '').localeCompare(a.registeredOn || '') ||
    (a.pageRank === undefined ? 999 : a.pageRank) - (b.pageRank === undefined ? 999 : b.pageRank);

  const sig = (o) => JSON.stringify([o.code, o.company, o.registeredOn]);

  function applyEntries(entries) {
    if (!entries.length) return;
    const data = load();
    const before = JSON.parse(JSON.stringify(data));
    let rank = 0;

    entries.forEach((e) => {
      const r = e.state === 'registered' ? rank++ : undefined;
      const old = data[e.code];

      if (e.state === 'not') { delete data[e.code]; return; }   // cancelled / not registered -> not tracked
      if (e.state === 'unknown') return;                          // leave whatever we already have

      data[e.code] = {
        code: e.code,
        company: e.company || (old && old.company) || '',
        registration: 'Registered',
        registeredOn: (old && old.registeredOn) || today(),
        pageRank: r,
      };
    });

    // Keep only the most recent N
    Object.values(data).sort(cmp).slice(CFG.maxDrives).forEach((d) => { delete data[d.code]; });

    if (JSON.stringify(data) !== JSON.stringify(before)) saveData(data);

    const kb = Object.keys(before), ka = Object.keys(data);
    const changed = kb.length !== ka.length || ka.some((k) => !before[k] || sig(before[k]) !== sig(data[k]));
    if (changed) GM_setValue(DIRTY, true);
  }

  /* -------------------- 1. live sync when on the registration page -------------------- */
  let syncTimer;
  const debouncedSync = () => {
    clearTimeout(syncTimer);
    syncTimer = setTimeout(() => applyEntries(readEntries(document)), CFG.syncDebounceMs);
  };

  /* -------------------- 2. background fetch from any other Placements page -------------------- */
  async function fetchAndSync() {
    GM_setValue(LAST_FETCH, Date.now());
    try {
      const res = await fetch(CFG.registrationPageUrl, { credentials: 'include' });
      if (!res.ok) return;
      const doc = new DOMParser().parseFromString(await res.text(), 'text/html');
      applyEntries(readEntries(doc));
    } catch (err) {
      console.warn('[UMS tracker] background fetch failed', err);
    }
  }

  function maybeAutoFetch() {
    if (window.top !== window.self || onRegistrationPage) return;
    if (Date.now() - GM_getValue(LAST_FETCH, 0) < CFG.autoFetchEveryMin * 60 * 1000) return;
    fetchAndSync();
  }

  // After you press a Register / Cancel link, refresh from the registration page once the portal updates
  let armed = false;
  document.addEventListener('click', (e) => {
    const a = e.target.closest('a');
    if (a && /lnkRegister$/.test(a.id)) {
      GM_setValue(LAST_FETCH, 0); // covers a full page reload
      armed = true;               // covers a partial (AJAX) update
    }
  }, true);

  new MutationObserver(() => {
    if (onRegistrationPage) debouncedSync();
    if (armed && !onRegistrationPage) {
      armed = false;
      setTimeout(fetchAndSync, 2000);
      setTimeout(fetchAndSync, 8000);
    }
  }).observe(document.documentElement, { childList: true, subtree: true });

  window.addEventListener('load', () => {
    if (onRegistrationPage) debouncedSync();
    else setTimeout(maybeAutoFetch, 2000);
  });

  /* --------------------------- 3. build + save the file --------------------------- */
  const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const HAS_PICKER = typeof window.showSaveFilePicker === 'function'; // Chrome / Edge only

  function buildBlob() {
    const rows = Object.values(load()).sort(cmp).map((d) => ({
      'Drive Name': d.company,
      'Drive Code': d.code,
      'Registration Status': d.registration || 'Registered',
      'Registered On (first seen)': d.registeredOn || '',
    }));
    if (!rows.length) return null;
    const ws = XLSX.utils.json_to_sheet(rows);
    ws['!cols'] = [34, 22, 20, 26].map((wch) => ({ wch }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Drives');
    return new Blob([XLSX.write(wb, { bookType: 'xlsx', type: 'array' })], { type: XLSX_MIME });
  }

  // Remember the chosen file so later saves overwrite it (Chrome / Edge)
  function idb() {
    return new Promise((res, rej) => {
      const r = indexedDB.open('ums_drive_tracker', 1);
      r.onupgradeneeded = () => r.result.createObjectStore('h');
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    });
  }
  async function getHandle() {
    try {
      const db = await idb();
      return await new Promise((res) => {
        const q = db.transaction('h').objectStore('h').get('file');
        q.onsuccess = () => res(q.result || null);
        q.onerror = () => res(null);
      });
    } catch (e) { return null; }
  }
  async function setHandle(h) {
    try {
      const db = await idb();
      await new Promise((res) => {
        const tx = db.transaction('h', 'readwrite');
        tx.objectStore('h').put(h, 'file');
        tx.oncomplete = () => res();
        tx.onerror = () => res();
      });
    } catch (e) { /* ignore */ }
  }

  async function writeToHandle(handle, blob) {
    const w = await handle.createWritable();
    await w.write(blob);
    await w.close();
  }

  async function saveWithPicker(blob, forcePick) {
    if (!forcePick) {
      const handle = await getHandle();
      if (handle) {
        try {
          if ((await handle.requestPermission({ mode: 'readwrite' })) === 'granted') {
            await writeToHandle(handle, blob);
            return;
          }
        } catch (e) { /* file moved or deleted -> ask again below */ }
      }
    }
    const handle = await window.showSaveFilePicker({
      suggestedName: CFG.fileName,
      types: [{ description: 'Excel workbook', accept: { [XLSX_MIME]: ['.xlsx'] } }],
    });
    await writeToHandle(handle, blob);
    await setHandle(handle);
  }

  // Firefox etc.: normal download. Turn on "Always ask you where to save files" to get a Save As dialog.
  function saveWithDownload(blob) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = CFG.fileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  }

  async function exportExcel(forcePick) {
    const blob = buildBlob();
    if (!blob) { alert('No registered drives recorded yet.'); return; }
    try {
      if (HAS_PICKER) await saveWithPicker(blob, forcePick === true);
      else saveWithDownload(blob);
      GM_setValue(DIRTY, false);
      hideBanner();
    } catch (err) {
      if (err && err.name === 'AbortError') return; // you closed the Save dialog
      console.error('[UMS tracker] save failed', err);
      alert('Could not save the file: ' + (err && err.message ? err.message : err));
    }
  }

  /* ----------------- 4. "changes made, save?" banner (top window) ----------------- */
  let banner;
  function makeBtn(label, bg, onClick) {
    const b = document.createElement('button');
    b.textContent = label;
    b.style.cssText = 'background:' + bg + ';border:1px solid #6b7280;color:#fff;padding:6px 12px;border-radius:6px;cursor:pointer';
    b.onclick = onClick;
    return b;
  }
  function showBanner() {
    if (banner || window.top !== window.self || !document.body) return;
    banner = document.createElement('div');
    banner.style.cssText =
      'position:fixed;bottom:16px;right:16px;z-index:2147483647;background:#1f2937;color:#fff;' +
      'padding:12px 16px;border-radius:8px;font:14px sans-serif;box-shadow:0 4px 12px rgba(0,0,0,.3);' +
      'display:flex;gap:10px;align-items:center';
    const msg = document.createElement('span');
    msg.textContent = 'Placement drive log updated.';
    banner.append(msg, makeBtn('Save Excel', '#22c55e', () => exportExcel(false)));
    if (HAS_PICKER) banner.append(makeBtn('Save As…', 'transparent', () => exportExcel(true)));
    banner.append(makeBtn('Later', 'transparent', hideBanner));
    document.body.appendChild(banner);
  }
  function hideBanner() { if (banner) { banner.remove(); banner = null; } }
  setInterval(() => { if (GM_getValue(DIRTY, false)) showBanner(); }, 2000);

  /* ----------------------------- menu commands ----------------------------- */
  GM_registerMenuCommand('Save drives to Excel', () => exportExcel(false));
  if (HAS_PICKER) GM_registerMenuCommand('Save As… (choose file)', () => exportExcel(true));
  GM_registerMenuCommand('Refresh from Drive Registration page now', () => {
    if (onRegistrationPage) applyEntries(readEntries(document)); else fetchAndSync();
  });
  GM_registerMenuCommand('Clear stored drives', () => {
    if (confirm('Delete all stored drive records?')) { saveData({}); GM_setValue(DIRTY, false); hideBanner(); }
  });
})();
