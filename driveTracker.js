// ==UserScript==
// @name         UMS Placement Drive Tracker
// @namespace    namit.ums.drive.tracker
// @version      5.0
// @description  Keeps your last 10 registered placement drives in an Excel file, with Conducted / Not conducted status per drive
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
    reportPageUrl: '/Placements/frmPlacementRegisteredStudentReport.aspx', // "Track Registered Drive Status"
    // Rounds matching this do NOT count as "conducted" (only the PPT has happened so far)
    pptRegex: /\b(ppt|pre[\s-]*placement[\s-]*talk|presentation)\b/i,
    autoRefreshEveryMin: 60,             // background refresh interval
    requestDelayMs: 400,                 // pause between portal requests
    fileName: 'placement_drives.xlsx',
    syncDebounceMs: 800,
  };

  const KEY = 'drives_v3';
  const DIRTY = 'dirty_v3';
  const LAST_REFRESH = 'last_refresh_v5';
  const onRegistrationPage = /frmPlacementDriveRegistration/i.test(location.pathname);

  /* ------------------------------ helpers ------------------------------ */
  const norm = (s) => (s || '').replace(/\s+/g, ' ').trim();
  const today = () => new Date().toISOString().slice(0, 10);
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const load = () => { try { return JSON.parse(GM_getValue(KEY, '{}')); } catch (e) { return {}; } };
  const saveData = (d) => GM_setValue(KEY, JSON.stringify(d));

  async function getDoc(url) {
    const res = await fetch(url, { credentials: 'include' });
    if (!res.ok) return null;
    return new DOMParser().parseFromString(await res.text(), 'text/html');
  }

  function toast(text) {
    if (window.top !== window.self || !document.body) return;
    const t = document.createElement('div');
    t.textContent = text;
    t.style.cssText = 'position:fixed;top:16px;right:16px;z-index:2147483647;background:#1f2937;color:#fff;' +
      'padding:10px 14px;border-radius:8px;font:14px sans-serif;box-shadow:0 4px 12px rgba(0,0,0,.3)';
    document.body.appendChild(t);
    setTimeout(() => t.remove(), 4000);
  }

  /* ------------------ registration page: which drives am I registered for? ------------------ */
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

  //   "Click to Cancel Registration" -> registered   (open drive, you're registered)
  //   "Click to Register"            -> not registered (open drive, you're not registered)
  //   "Yes"                          -> registered   (closed drive - registration closed, plain Yes/No text)
  //   "No"                           -> not registered
  //   anything else                  -> unknown (stored drives are left untouched)
  function registrationState(text) {
    if (/cancel/i.test(text)) return 'registered';
    if (/click to register/i.test(text)) return 'not';
    const t = text.trim().toLowerCase();
    if (t === 'yes') return 'registered';
    if (t === 'no') return 'not';
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

  const sig = (o) => JSON.stringify([o.code, o.company, o.registeredOn, o.driveStatus]);

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

      const rec = {
        code: e.code,
        company: e.company || (old && old.company) || '',
        registration: 'Registered',
        registeredOn: (old && old.registeredOn) || today(),
        pageRank: r,
      };
      if (old && old.driveStatus) rec.driveStatus = old.driveStatus;
      data[e.code] = rec;
    });

    // Keep only the most recent N
    Object.values(data).sort(cmp).slice(CFG.maxDrives).forEach((d) => { delete data[d.code]; });

    if (JSON.stringify(data) !== JSON.stringify(before)) saveData(data);

    const kb = Object.keys(before), ka = Object.keys(data);
    const changed = kb.length !== ka.length || ka.some((k) => !before[k] || sig(before[k]) !== sig(data[k]));
    if (changed) GM_setValue(DIRTY, true);
  }

  async function syncFromRegistrationPage() {
    try {
      const doc = await getDoc(CFG.registrationPageUrl);
      if (doc) applyEntries(readEntries(doc));
    } catch (err) {
      console.warn('[UMS tracker] registration fetch failed', err);
    }
  }

  /* ------------- report page ("Track Registered Drive Status"): Conducted / Not conducted ------------- */
  // Replays the page's own ASP.NET postbacks in the background (no navigation, nothing visible)
  async function postBack(doc, eventTarget, fields) {
    const form = doc.querySelector('form');
    if (!form) return null;
    const body = new URLSearchParams();
    form.querySelectorAll('input, select, textarea').forEach((el) => {
      if (!el.name || el.disabled) return;
      const type = (el.type || '').toLowerCase();
      if (['submit', 'button', 'image', 'file', 'reset'].indexOf(type) >= 0) return;
      if ((type === 'checkbox' || type === 'radio') && !el.checked) return;
      body.set(el.name, el.value);
    });
    body.set('__EVENTTARGET', eventTarget || '');
    body.set('__EVENTARGUMENT', '');
    Object.keys(fields || {}).forEach((k) => body.set(k, fields[k]));

    const res = await fetch(CFG.reportPageUrl, {
      method: 'POST',
      credentials: 'include',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: body.toString(),
    });
    if (!res.ok) return null;
    return new DOMParser().parseFromString(await res.text(), 'text/html');
  }

  // Dates on the portal look like 09/17/2026 (MM/DD/YYYY). A round dated in the future is not conducted yet.
  function isFuture(s) {
    const m = (s || '').match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (!m) return false;
    const a = +m[1], b = +m[2], y = +m[3];
    const mo = a > 12 ? b : a, d = a > 12 ? a : b;
    const end = new Date(); end.setHours(23, 59, 59, 999);
    return new Date(y, mo - 1, d) > end;
  }

  const findDriveRow = (doc, code) => {
    const link = Array.from(doc.querySelectorAll('a[id$="hypDetails"]')).find((a) => norm(a.textContent) === code);
    return link ? link.closest('tr') : null;
  };

  const totalPages = (doc) => {
    const info = doc.querySelector('.rgInfoPart');
    const m = info && norm(info.textContent).match(/items in\s*(\d+)\s*pages?/i);
    return m ? +m[1] : 1;
  };

  // Returns 'Conducted', 'Not conducted', or null if the portal response could not be understood.
  // The report only lists drives where at least one round has happened, so a drive that is not listed = Not conducted.
  async function fetchDriveStatus(baseDoc, drive) {
    const filterInput = baseDoc.querySelector('input[id$="txtFilter"]');
    if (!filterInput) return null;
    const fname = filterInput.name;

    // 1. narrow the list to this company (the report is paged, oldest drives first, newest on the last pages)
    let page = await postBack(baseDoc, fname, { [fname]: drive.company });
    if (!page) return null;

    // make sure the portal really processed the filter and returned the grid, otherwise "not listed" would be meaningless
    const echo = page.querySelector('input[id$="txtFilter"]');
    if (!page.querySelector('table[id$="grddata_ctl00"]') || !echo ||
        norm(echo.value).toLowerCase() !== drive.company.toLowerCase()) return null;

    // 2. find the exact Drive Code (look through further pages of the filtered list if there are any)
    let row = findDriveRow(page, drive.code);
    const pages = Math.min(totalPages(page), 5);
    for (let p = 2; !row && p <= pages; p++) {
      const a = Array.from(page.querySelectorAll('.rgNumPart a')).find((x) => norm(x.textContent) === String(p));
      const m = a && (a.getAttribute('href') || '').match(/__doPostBack\('([^']+)'/);
      if (!m) return null; // cannot reach the next page -> do not guess
      await sleep(CFG.requestDelayMs);
      page = await postBack(page, m[1], { [fname]: drive.company });
      if (!page) return null;
      row = findDriveRow(page, drive.code);
    }

    // not listed at all -> no round has been conducted for this drive
    if (!row) return 'Not conducted';

    const btn = row.querySelector('input[id$="btnView"]');
    if (!btn) return null;
    const rowCompanyEl = row.querySelector('span[id$="lblCompanyName"]');
    const rowCompany = rowCompanyEl ? norm(rowCompanyEl.textContent).toUpperCase() : '';

    // 3. press its "Drive Status" button
    await sleep(CFG.requestDelayMs);
    const result = await postBack(page, '', { [fname]: drive.company, [btn.name]: btn.value || 'Drive Status' });
    if (!result) return null;

    // 4. read the rounds table shown in the "Drive Details" popup
    const grid = result.querySelector('table[id$="GridView1_ctl00"]') || result.querySelector('[id$="GridView1"]');
    if (!grid) return null;
    const rounds = Array.from(grid.querySelectorAll('span[id$="lblRoundName"]')).map((s) => {
      const tr = s.closest('tr');
      const q = (sel) => { const el = tr && tr.querySelector(sel); return el ? norm(el.textContent) : ''; };
      return { name: norm(s.textContent), on: q('span[id$="lblConductedOn"]'), company: q('span[id$="lblCompanyName"]').toUpperCase() };
    });

    // sanity check: the rounds shown must belong to the company we clicked
    if (rounds.length && rowCompany && !rounds.some((r) => r.company === rowCompany)) return null;

    const conducted = rounds.some((r) => r.name && !CFG.pptRegex.test(r.name) && !isFuture(r.on));
    return conducted ? 'Conducted' : 'Not conducted';
  }

  function setStatus(code, status) {
    const data = load();
    const d = data[code];
    if (!d || d.driveStatus === status) return;
    d.driveStatus = status;
    saveData(data);
    GM_setValue(DIRTY, true);
  }

  let checking = false;
  async function checkStatuses() {
    if (checking) return;
    checking = true;
    try {
      // "Conducted" can never go back, so only drives not yet conducted are re-checked
      const todo = Object.values(load()).filter((d) => d.driveStatus !== 'Conducted');
      if (!todo.length) return;
      const baseDoc = await getDoc(CFG.reportPageUrl);
      if (!baseDoc) return;
      for (const d of todo) {
        try {
          const status = await fetchDriveStatus(baseDoc, d);
          if (status) setStatus(d.code, status);
        } catch (err) {
          console.warn('[UMS tracker] status check failed for', d.code, err);
        }
        await sleep(CFG.requestDelayMs);
      }
    } catch (err) {
      console.warn('[UMS tracker] status check failed', err);
    } finally {
      checking = false;
    }
  }

  /* ------------------------------ when things run ------------------------------ */
  // Live sync when you are on the registration page
  let syncTimer;
  const debouncedSync = () => {
    clearTimeout(syncTimer);
    syncTimer = setTimeout(() => applyEntries(readEntries(document)), CFG.syncDebounceMs);
  };

  async function refreshAll() {
    if (!onRegistrationPage) await syncFromRegistrationPage();
    await checkStatuses();
  }

  function maybeRefresh() {
    if (window.top !== window.self) return;
    if (Date.now() - GM_getValue(LAST_REFRESH, 0) < CFG.autoRefreshEveryMin * 60 * 1000) return;
    GM_setValue(LAST_REFRESH, Date.now());
    refreshAll();
  }

  // After you press a Register / Cancel link, refresh once the portal has updated
  let armed = false;
  document.addEventListener('click', (e) => {
    const a = e.target.closest('a');
    if (a && /lnkRegister$/.test(a.id)) {
      GM_setValue(LAST_REFRESH, 0); // covers a full page reload
      armed = true;                 // covers a partial (AJAX) update
    }
  }, true);

  new MutationObserver(() => {
    if (onRegistrationPage) debouncedSync();
    if (armed && !onRegistrationPage) {
      armed = false;
      setTimeout(syncFromRegistrationPage, 2000);
      setTimeout(async () => { await syncFromRegistrationPage(); await checkStatuses(); }, 8000);
    }
  }).observe(document.documentElement, { childList: true, subtree: true });

  window.addEventListener('load', () => {
    if (onRegistrationPage) debouncedSync();
    setTimeout(maybeRefresh, 2500);
  });

  /* --------------------------- build + save the file --------------------------- */
  const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const HAS_PICKER = typeof window.showSaveFilePicker === 'function'; // Chrome / Edge only

  function buildBlob() {
    const rows = Object.values(load()).sort(cmp).map((d) => ({
      'Drive Name': d.company,
      'Drive Code': d.code,
      'Registration Status': d.registration || 'Registered',
      'Registered On (first seen)': d.registeredOn || '',
      'Drive Status': d.driveStatus || 'Not checked',
    }));
    if (!rows.length) return null;
    const ws = XLSX.utils.json_to_sheet(rows);
    ws['!cols'] = [34, 22, 20, 26, 16].map((wch) => ({ wch }));
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

  /* ----------------- "changes made, save?" banner (top window) ----------------- */
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
  GM_registerMenuCommand('Refresh registrations + drive status now', async () => {
    toast('Refreshing drive data…');
    if (onRegistrationPage) applyEntries(readEntries(document));
    await refreshAll();
    toast('Refresh finished');
  });
  GM_registerMenuCommand('Clear stored drives', () => {
    if (confirm('Delete all stored drive records?')) { saveData({}); GM_setValue(DIRTY, false); hideBanner(); }
  });
})();
