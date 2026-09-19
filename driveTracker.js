// ==UserScript==
// @name         UMS Placement Drive Tracker
// @namespace    namit.ums.drive.tracker
// @version      3.0
// @description  Keeps the last 10 registered placement drives (from the UMS Drive Registration page) in an Excel file
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
    maxDrives: 10,                       // how many registered drives to keep in the file
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
  const stamp = () => new Date().toLocaleString('en-IN');
  const load = () => { try { return JSON.parse(GM_getValue(KEY, '{}')); } catch (e) { return {}; } };
  const saveData = (d) => GM_setValue(KEY, JSON.stringify(d));

  // Drive Registration page columns (the home-page table has no "Drive Code", so it is ignored)
  const COLS = {
    code: /^drive code/i,
    driveDate: /^drive date/i,
    registerBy: /^register by/i,
    company: /^company/i,
    venue: /^venue/i,
    driveStatus: /^status$/i,
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

  // Decide registration state from the text of the "Registered" cell
  //   "Click to Cancel Registration" -> registered
  //   "Click to Register"            -> not registered
  //   anything else                  -> unknown (existing status is left untouched)
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
        const t = (i) => (i === undefined ? '' : norm(c[i].textContent));
        const code = t(map.code);
        if (!code) return;
        const portalText = t(map.registered);
        entries.push({
          code,
          company: t(map.company),
          registerBy: t(map.registerBy),
          driveDate: t(map.driveDate),
          venue: t(map.venue),
          driveStatus: t(map.driveStatus),
          portalText,
          state: registrationState(portalText),
        });
      });
    });
    return entries;
  }

  // Registered first, then newest registration date, then position on the portal page (top = newest)
  const cmp = (a, b) =>
    Number(b.registration === 'Registered') - Number(a.registration === 'Registered') ||
    (b.registeredOn || '').localeCompare(a.registeredOn || '') ||
    (a.pageRank === undefined ? 999 : a.pageRank) - (b.pageRank === undefined ? 999 : b.pageRank);

  const strip = (o) => { const c = Object.assign({}, o); delete c.lastUpdated; delete c.pageRank; return JSON.stringify(c); };

  function applyEntries(entries) {
    if (!entries.length) return;
    const data = load();
    const before = JSON.parse(JSON.stringify(data));
    let rank = 0;

    entries.forEach((e) => {
      const r = e.state === 'registered' ? rank++ : undefined;
      const old = data[e.code];

      let registration;
      if (e.state === 'registered') registration = 'Registered';
      else if (e.state === 'not' && old) registration = 'Not registered';
      else if (e.state === 'unknown' && old) registration = old.registration;
      else return; // not registered and not tracked -> ignore

      const fields = { code: e.code, company: e.company, registerBy: e.registerBy, driveDate: e.driveDate,
                       venue: e.venue, driveStatus: e.driveStatus, portalText: e.portalText };
      Object.keys(fields).forEach((k) => { if (!fields[k]) delete fields[k]; });

      const next = Object.assign({}, old || {}, fields, { registration });
      if (registration === 'Registered') {
        next.registeredOn = (old && old.registeredOn) || today();
        if (r !== undefined) next.pageRank = r;
      } else if (registration === 'Not registered') {
        next.registeredOn = '';
      }
      next.lastUpdated = !old || strip(old) !== strip(next) ? stamp() : old.lastUpdated;
      data[e.code] = next;
    });

    // Keep only the most recent N
    Object.values(data).sort(cmp).slice(CFG.maxDrives).forEach((d) => { delete data[d.code]; });

    if (JSON.stringify(data) !== JSON.stringify(before)) saveData(data);

    const kb = Object.keys(before), ka = Object.keys(data);
    const changed = kb.length !== ka.length || ka.some((k) => !before[k] || strip(before[k]) !== strip(data[k]));
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

  /* ------------------------------ 3. export ------------------------------ */
  function exportExcel() {
    const rows = Object.values(load()).sort(cmp).map((d) => ({
      'Company': d.company,
      'Drive Code': d.code,
      'Register By': d.registerBy || '',
      'Drive Date': d.driveDate || '',
      'Venue': d.venue || '',
      'Drive Status': d.driveStatus || '',
      'Registration Status': d.registration || '',
      'Registered On (first seen)': d.registeredOn || '',
      'Last Updated': d.lastUpdated || '',
      'Portal Text': d.portalText || '',
    }));
    if (!rows.length) { alert('No registered drives recorded yet.'); return; }
    const ws = XLSX.utils.json_to_sheet(rows);
    ws['!cols'] = [30, 22, 18, 22, 26, 12, 20, 24, 22, 30].map((wch) => ({ wch }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Drives');
    XLSX.writeFile(wb, CFG.fileName);
    GM_setValue(DIRTY, false);
    hideBanner();
  }

  /* ----------------- 4. "changes made, save?" banner (top window) ----------------- */
  let banner;
  function showBanner() {
    if (banner || window.top !== window.self || !document.body) return;
    banner = document.createElement('div');
    banner.style.cssText =
      'position:fixed;bottom:16px;right:16px;z-index:2147483647;background:#1f2937;color:#fff;' +
      'padding:12px 16px;border-radius:8px;font:14px sans-serif;box-shadow:0 4px 12px rgba(0,0,0,.3);' +
      'display:flex;gap:10px;align-items:center';
    banner.innerHTML = '<span>Placement drive log updated.</span>';
    const saveBtn = document.createElement('button');
    saveBtn.textContent = 'Save Excel';
    saveBtn.style.cssText = 'background:#22c55e;border:0;color:#fff;padding:6px 12px;border-radius:6px;cursor:pointer';
    saveBtn.onclick = exportExcel;
    const later = document.createElement('button');
    later.textContent = 'Later';
    later.style.cssText = 'background:transparent;border:1px solid #6b7280;color:#fff;padding:6px 10px;border-radius:6px;cursor:pointer';
    later.onclick = hideBanner;
    banner.append(saveBtn, later);
    document.body.appendChild(banner);
  }
  function hideBanner() { if (banner) { banner.remove(); banner = null; } }
  setInterval(() => { if (GM_getValue(DIRTY, false)) showBanner(); }, 2000);

  /* ----------------------------- menu commands ----------------------------- */
  GM_registerMenuCommand('Export drives to Excel', exportExcel);
  GM_registerMenuCommand('Refresh from Drive Registration page now', () => {
    if (onRegistrationPage) applyEntries(readEntries(document)); else fetchAndSync();
  });
  GM_registerMenuCommand('Clear stored drives', () => {
    if (confirm('Delete all stored drive records?')) { saveData({}); GM_setValue(DIRTY, false); hideBanner(); }
  });
})();
