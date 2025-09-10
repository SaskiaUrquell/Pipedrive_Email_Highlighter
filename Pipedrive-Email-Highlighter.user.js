// ==UserScript==
// @name         Pipedrive Email Highlighter 
// @namespace    https://pakajo-helper
// @version      1.7.7
// @description  Markiert E-Mails rot/gelb/grün … 
// @match        http*://*/*
// @grant        GM_xmlhttpRequest
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_registerMenuCommand
// @grant        GM_addStyle
// @connect      api.pipedrive.com
// @run-at       document-end
// @noframes
// @updateURL    https://raw.githubusercontent.com/SaskiaUrquell/Pipedrive_Email_Highlighter/main/Pipedrive-Email-Highlighter.user.js
// @downloadURL  https://raw.githubusercontent.com/SaskiaUrquell/Pipedrive_Email_Highlighter/main/Pipedrive-Email-Highlighter.user.js
// ==/UserScript==


(function () {
  // ===== Einstellungen =====
  const CHECK_LEADS = true;
  const THROTTLE_MS = 250;

  // Request-Policy
  const REQUEST_TIMEOUT_MS = 10000; // 10s Timeout je Request
  const MAX_RETRIES = 2;            // +2 Wiederholungen
  const RETRY_BASE_DELAY_MS = 800;  // 0.8s, dann 1.6s Backoff
  const DEEP_ORG_DETAIL_LIMIT = 6;  // max. Orga-Detailabrufe pro E-Mail

  // Sichtbarkeits-/Scope-Optionen (per Menü änderbar)
  const VISIBLE_ONLY_DEFAULT  = true; // nur aktiver Tab
  const VIEWPORT_ONLY_DEFAULT = true; // nur sichtbarer Bereich

  // ===== CSS =====
  GM_addStyle(`
    .pd-email-red    { background: rgba(255,0,0,.15) !important; outline: 2px solid rgba(255,0,0,.6) !important; border-radius:4px !important; padding:0 3px !important; }
    .pd-email-yellow { background: rgba(255,200,0,.18) !important; outline: 2px solid rgba(255,200,0,.7) !important; border-radius:4px !important; padding:0 3px !important; }
    .pd-email-green  { background: rgba(0,200,0,.12) !important; outline: 2px solid rgba(0,200,0,.5) !important; border-radius:4px !important; padding:0 3px !important; }
    .pd-email-error  { background: rgba(128,128,128,.12) !important; outline: 2px dashed rgba(128,128,128,.6) !important; border-radius:4px !important; padding:0 3px !important; }
  `);

  // ===== Konstanten & Utils =====
  const API_V1 = 'https://api.pipedrive.com/v1';
  const API_V2 = 'https://api.pipedrive.com/api/v2';
  const EMAIL_RX = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi;
  const sleep = (ms) => new Promise(r => setTimeout(r, ms));
  const norm = (s) => String(s || '').toLowerCase().trim();
  const normEmail = (s) => norm(String(s).replace(/^mailto:/i, ''));

  // Obfuskations-Tokens
  const OB_AT  = '(?:@|\\(at\\)|\\[at\\]|\\{at\\}|&commat;|\\bat\\b|\\bät\\b)';
  const OB_DOT = '(?:\\.|\\(dot\\)|\\[dot\\]|\\{dot\\}|\\bdot\\b|\\bpunkt\\b)';
  const OBFUSCATED_RX = new RegExp(
    '([A-Z0-9._%+-]+)\\s*' + OB_AT + '\\s*([A-Z0-9][A-Z0-9.-]*(?:\\s*' + OB_DOT + '\\s*[A-Z0-9][A-Z0-9.-]*)*)',
    'gi'
  );
  const DOT_TOKEN_RX = new RegExp(OB_DOT, 'gi');

  // ===== Ein/Aus =====
  function isEnabled() { return GM_getValue('pd_enabled', true); }
  function setEnabled(v) { GM_setValue('pd_enabled', !!v); }

  function getVisibleOnly()  { return GM_getValue('pd_visible_only',  VISIBLE_ONLY_DEFAULT); }
  function setVisibleOnly(v) { GM_setValue('pd_visible_only', !!v); }
  function getViewportOnly() { return GM_getValue('pd_viewport_only', VIEWPORT_ONLY_DEFAULT); }
  function setViewportOnly(v){ GM_setValue('pd_viewport_only', !!v); }

  let observer = null;
  function stopObserver() { if (observer) { observer.disconnect(); observer = null; } }
  function startObserver() {
    if (!document.body || observer || !isEnabled()) return;
    if (getVisibleOnly() && document.hidden) return; // nur aktiver Tab
    observer = new MutationObserver(() => {
      scheduleScan(150);
    });
    observer.observe(document.body, { childList: true, subtree: true });
  }
  function toggleEnabled() {
    const now = !isEnabled();
    setEnabled(now);
    if (now) { startObserver(); scan(); alert('Pipedrive-Markierung: AN'); }
    else { stopObserver(); alert('Pipedrive-Markierung: AUS'); }
  }
  function toggleVisibleOnly(){
    setVisibleOnly(!getVisibleOnly());
    alert('Nur aktiver Tab: ' + (getVisibleOnly() ? 'AN' : 'AUS'));
    // Sofort anwenden
    stopObserver(); startObserver(); scheduleScan(50);
  }
  function toggleViewportOnly(){
    setViewportOnly(!getViewportOnly());
    alert('Nur sichtbarer Bereich: ' + (getViewportOnly() ? 'AN' : 'AUS'));
    scheduleScan(50);
  }

  // ===== Token =====
  function getToken() {
    let token = GM_getValue('pd_token', '');
    if (!token) {
      token = prompt('[Pipedrive] Bitte Deinen persönlichen API-Token einfügen (App → Personal preferences → API).');
      if (token) GM_setValue('pd_token', token.trim());
    }
    return token || '';
  }
  function setToken() { const t = prompt('Neuen Pipedrive API-Token eingeben:'); if (t) GM_setValue('pd_token', t.trim()); }
  function clearToken() { GM_setValue('pd_token', ''); alert('Pipedrive-Token gelöscht.'); }

  // ===== Menü =====
  GM_registerMenuCommand(`Markierung AN/AUS (Pipedrive) — v${SCRIPT_VERSION}`, toggleEnabled);
  GM_registerMenuCommand('Nur aktiver Tab (AN/AUS)', toggleVisibleOnly);
  GM_registerMenuCommand('Nur sichtbarer Bereich (AN/AUS)', toggleViewportOnly);
  GM_registerMenuCommand('Pipedrive-Token setzen/ändern', setToken);
  GM_registerMenuCommand('Pipedrive-Token löschen', clearToken);
  GM_registerMenuCommand('Seite scannen (Pipedrive)', () => scan(true));

  // ===== Persistente Caches (über Tabs/Fenster) =====
  const emailCacheObj  = safeParse(GM_getValue('pd_email_cache',  '{}'));
  const domainCacheObj = safeParse(GM_getValue('pd_domain_cache', '{}'));
  const cache = new Map(Object.entries(emailCacheObj));        // email -> status
  const domainCache = new Map(Object.entries(domainCacheObj)); // domain -> boolean
  let persistTimer = null;
  function safeParse(s) { try { return JSON.parse(s || '{}'); } catch { return {}; } }
  function schedulePersist() { if (!persistTimer) persistTimer = setTimeout(persistNow, 30000); }
  function persistNow() {
    if (persistTimer) { clearTimeout(persistTimer); persistTimer = null; }
    try {
      GM_setValue('pd_email_cache',  JSON.stringify(Object.fromEntries(cache)));
      GM_setValue('pd_domain_cache', JSON.stringify(Object.fromEntries(domainCache)));
    } catch {}
  }

  // ===== HTTP mit Timeout & Retry =====
  function parseRetryAfter(headers) {
    try { const m = /retry-after:\s*([0-9]+)/i.exec(headers || ''); return m ? Number(m[1]) : 0; } catch { return 0; }
  }
  function req(base, path, attempt = 0) {
    const token = getToken();
    if (!token) throw new Error('Kein Pipedrive-Token gesetzt.');
    const url = `${base}${path}${path.includes('?') ? '&' : '?'}api_token=${encodeURIComponent(token)}`;
    return new Promise((resolve, reject) => {
      GM_xmlhttpRequest({
        method: 'GET',
        url,
        timeout: REQUEST_TIMEOUT_MS,
        onload: (resp) => {
          try {
            if (resp.status === 401) return reject(new Error('401 Unauthorized'));
            if (resp.status === 429) {
              if (attempt < MAX_RETRIES) {
                const ra = parseRetryAfter(resp.responseHeaders);
                const delay = ra ? ra * 1000 : RETRY_BASE_DELAY_MS * Math.pow(2, attempt);
                return setTimeout(() => req(base, path, attempt + 1).then(resolve, reject), delay);
              }
              return reject(new Error('429 Rate Limit'));
            }
            if (resp.status < 200 || resp.status >= 300) return reject(new Error('HTTP ' + resp.status));
            resolve(JSON.parse(resp.responseText));
          } catch (e) { reject(e); }
        },
        ontimeout: () => {
          if (attempt < MAX_RETRIES) {
            const delay = RETRY_BASE_DELAY_MS * Math.pow(2, attempt);
            setTimeout(() => req(base, path, attempt + 1).then(resolve, reject), delay);
          } else reject(new Error('TIMEOUT'));
        },
        onerror: () => reject(new Error('NETWORK'))
      });
    });
  }
  const requestV1 = (p) => req(API_V1, p);
  const requestV2 = (p) => req(API_V2, p);

  // ===== Helper: E-Mails aus Suchitems ziehen =====
  function extractEmailsFromItem(it) {
    const out = [];
    const push = (v) => { if (v) out.push(normEmail(v)); };
    const emails1 = it?.item?.email_addresses || [];
    for (const x of emails1) push(x?.value ?? x);
    const emails2 = it?.item?.emails || [];
    for (const x of emails2) push(x?.value ?? x);
    push(it?.item?.primary_email);
    const fm = it?.field_matches || it?.matches || it?.item?.matches || [];
    const arr = Array.isArray(fm) ? fm : [fm];
    for (const m of arr) push(m?.content ?? m);
    return out.filter(Boolean);
  }
  function emailDomain(e) { const at = e.lastIndexOf('@'); return at >= 0 ? e.slice(at + 1) : ''; }
  function matchesDomain(emailLower, d) { const ed = emailDomain(emailLower); return ed === d || ed.endsWith('.' + d); }

  // ===== Personen/Leads/Organisationen =====
  async function existsPersonByEmail(email) {
    const e = normEmail(email);
    try { const r = await requestV2(`/persons/search?term=${encodeURIComponent(e)}&fields=email&exact_match=1&limit=1`); if ((r?.data?.items || []).length > 0) return true; } catch {}
    try { const r = await requestV1(`/persons/find?term=${encodeURIComponent(e)}&search_by_email=1&limit=1`); if ((r?.data || []).length > 0) return true; } catch {}
    return false;
  }
  async function existsLeadByEmail(email) {
    const e = normEmail(email);
    try { const r = await requestV2(`/itemSearch?term=${encodeURIComponent(e)}&item_types=lead&fields=email&exact_match=1&limit=1`); return (r?.data?.items || []).length > 0; } catch {}
    return false;
  }
  function objectContainsEmail(obj, targetEmail) {
    const target = normEmail(targetEmail);
    const stack = [obj]; const seen = new Set();
    while (stack.length) {
      const cur = stack.pop();
      if (!cur || typeof cur !== 'object') continue;
      if (seen.has(cur)) continue; seen.add(cur);
      for (const k of Object.keys(cur)) {
        const v = cur[k];
        if (typeof v === 'string') {
          if (normEmail(v) === target) return true;
          EMAIL_RX.lastIndex = 0;
          const ms = v.match(EMAIL_RX);
          if (ms && ms.map(normEmail).includes(target)) return true;
        } else if (v && typeof v === 'object') stack.push(v);
      }
    }
    return false;
  }
  async function existsOrgByEmail(email) {
    const e = normEmail(email);
    const candidateIds = new Set();
    const paths = [
      `/itemSearch?term=${encodeURIComponent(e)}&item_types=organization&fields=email,custom_fields&exact_match=1&limit=50`,
      `/itemSearch?term=${encodeURIComponent(e)}&item_types=organization&exact_match=1&limit=50`,
      `/itemSearch?term=${encodeURIComponent(e)}&item_types=organization&limit=50`
    ];
    for (const p of paths) {
      try {
        const res = await requestV2(p);
        const items = res?.data?.items || [];
        for (const it of items) {
          const emails = extractEmailsFromItem(it);
          if (emails.includes(e)) return true;
          const id = it?.item?.id ?? it?.id;
          if (id != null) candidateIds.add(id);
        }
      } catch {}
    }
    try {
      const res = await requestV1(`/organizations/search?term=${encodeURIComponent(e)}&limit=50`);
      const items = res?.data?.items || res?.data || [];
      for (const it of items) { const id = it?.item?.id ?? it?.id; if (id != null) candidateIds.add(id); }
    } catch {}

    const ids = Array.from(candidateIds).slice(0, DEEP_ORG_DETAIL_LIMIT);
    for (const id of ids) {
      try {
        const det = await requestV1(`/organizations/${id}`);
        const org = det?.data || det;
        if (org && objectContainsEmail(org, e)) return true;
      } catch {}
      await sleep(THROTTLE_MS);
    }
    return false;
  }

  // ===== Gelb: Domain-Indiz (Personen & Organisationen, inkl. Subdomains) =====
  const domainInflight = new Map(); // d -> Promise<boolean>
  function emailsFromItems(items) { const all = []; for (const it of (items || [])) all.push(...extractEmailsFromItem(it)); return all; }
  async function anyContactOrOrgWithDomain(domain) {
    const d = norm(domain);
    if (!d) return false;
    if (domainCache.has(d)) return domainCache.get(d);
    if (domainInflight.has(d)) return domainInflight.get(d);

    const p = (async () => {
      const paths = [
        `/persons/search?term=${encodeURIComponent('@' + d)}&limit=50`,
        `/persons/search?term=${encodeURIComponent(d)}&limit=50`,
        `/itemSearch?term=${encodeURIComponent('@' + d)}&item_types=person&fields=email&limit=50`,
        `/itemSearch?term=${encodeURIComponent('@' + d)}&item_types=organization&fields=email,custom_fields&limit=50`,
        `/itemSearch?term=${encodeURIComponent(d)}&item_types=organization&limit=50`
      ];
      for (const path of paths) {
        try {
          const res = await requestV2(path);
          const emails = emailsFromItems(res?.data?.items || []);
          if (emails.some(e => matchesDomain(e, d))) { domainCache.set(d, true); schedulePersist(); return true; }
        } catch {}
      }
      domainCache.set(d, false); schedulePersist();
      return false;
    })();

    domainInflight.set(d, p);
    try { return await p; } finally { domainInflight.delete(d); }
  }

  // ===== Status-Logik =====
  async function statusForEmail(email) {
    const e = normEmail(email);
    if (!e || !e.includes('@')) return 'green';

    // Rot: exakter Treffer
    if (await existsPersonByEmail(e)) return 'red';
    if (CHECK_LEADS) { try { if (await existsLeadByEmail(e)) return 'red'; } catch {} }
    if (await existsOrgByEmail(e)) return 'red';

    // Gelb: Domain-Indiz
    const d = emailDomain(e);
    if (d && await anyContactOrOrgWithDomain(d)) return 'yellow';

    // Grün: nix gefunden
    return 'green';
  }

  // ===== No-Rescan: Cache + parallele Entdoppelung =====
  const inflight = new Map(); // email -> Promise<status>
  async function getStatusCached(email) {
    const key = normEmail(email);
    if (cache.has(key)) return cache.get(key);
    if (inflight.has(key)) return inflight.get(key);
    const p = (async () => {
      try {
        const st = await statusForEmail(key);
        cache.set(key, st); schedulePersist();
        return st;
      } catch {
        cache.set(key, 'error'); schedulePersist();
        return 'error';
      } finally {
        inflight.delete(key);
      }
    })();
    inflight.set(key, p);
    return p;
  }

  function applyStyle(el, st) {
    el.classList.remove('pd-email-red','pd-email-yellow','pd-email-green','pd-email-error');
    const cls = st === 'red' ? 'pd-email-red' : st === 'yellow' ? 'pd-email-yellow' : st === 'green' ? 'pd-email-green' : 'pd-email-error';
    el.classList.add(cls);
    const label = st === 'red' ? 'existiert (Person/Lead/Organisation)' :
                  st === 'yellow' ? 'mgl. vorhanden (gleiche Domain)' :
                  st === 'green' ? 'nicht gefunden' : 'Fehler';
    el.title = `Pipedrive: ${label}`;
  }

  // Sichtbarkeit/Viewport
  function isInViewport(el) {
    try {
      const r = el.getBoundingClientRect();
      if (!r || r.width === 0 || r.height === 0) return false;
      const vw = window.innerWidth || document.documentElement.clientWidth;
      const vh = window.innerHeight || document.documentElement.clientHeight;
      return r.bottom >= 0 && r.right >= 0 && r.top <= vh && r.left <= vw;
    } catch { return true; }
  }

  // ===== DOM-Scan =====
  function shouldSkipTextNode(node) {
    let el = node.parentElement;
    while (el) {
      const tn = el.tagName || '';
      if (tn === 'A') return true; // bereits Link
      if (el.isContentEditable) return true;
      if (el.matches && el.matches(
        'input, textarea, select, option, button,' +
        '[contenteditable=""], [contenteditable="true"],' +
        '[role="textbox"], [role="combobox"], [role="listbox"],' +
        '[data-pd-skip], script, style, code, pre, kbd, samp, svg, canvas'
      )) return true;
      el = el.parentElement;
    }
    return false;
  }

  function linkEmailsInTextNode(node) {
    const parentEl = node.parentElement;
    if (!parentEl) return;
    if (getViewportOnly() && parentEl && !isInViewport(parentEl)) return; // nur sichtbarer Bereich
    if (shouldSkipTextNode(node)) return;

    const text = node.nodeValue || '';
    const matches = [];

    // 1) Normale E-Mails
    EMAIL_RX.lastIndex = 0;
    let m;
    while ((m = EMAIL_RX.exec(text))) {
      matches.push({ start: m.index, end: m.index + m[0].length, email: m[0] });
    }

    // 2) Obfuskierte E-Mails (info[at]domain dot com, etc.)
    OBFUSCATED_RX.lastIndex = 0;
    while ((m = OBFUSCATED_RX.exec(text))) {
      const user = m[1];
      const rawDomain = m[2];
      const domain = rawDomain
        .replace(DOT_TOKEN_RX, '.')   // dot/punkt → .
        .replace(/[()\[\]{}]/g, '')   // Klammern raus
        .replace(/\s+/g, '');         // Leerzeichen raus
      const email = `${user}@${domain}`;
      if (/^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$/i.test(email)) {
        matches.push({ start: m.index, end: m.index + m[0].length, email });
      }
    }

    if (!matches.length) return;

    // Nach Start index sortieren und Überlappungen vermeiden
    matches.sort((a,b) => a.start - b.start);
    const frag = document.createDocumentFragment();
    let last = 0;
    for (const mt of matches) {
      if (mt.start < last) continue; // überlappt bereits ersetztes Stück
      frag.append(text.slice(last, mt.start));
      const a = document.createElement('a');
      a.href = 'mailto:' + mt.email;
      a.textContent = mt.email;
      a.dataset._pd_plain = '1';
      frag.append(a);
      last = mt.end;
    }
    frag.append(text.slice(last));
    parentEl.replaceChild(frag, node);
  }

  async function processMailtos(root = document) {
    if (!isEnabled()) return;
    if (getVisibleOnly() && document.hidden) return;

    const anchors = Array.from(root.querySelectorAll('a')).filter(a => {
      const href = norm(a.getAttribute('href') || '');
      if (!href.startsWith('mailto:')) return false;
      return getViewportOnly() ? isInViewport(a) : true;
    });

    for (const a of anchors) {
      if (a.dataset._pd_done === '1') continue;
      a.dataset._pd_done = '1'; // gegen Parallel-Scans

      const raw = (a.getAttribute('href') || '').replace(/^mailto:/i, '').split('?')[0].trim();
      const emails = raw.split(',').map(s => s.trim()).filter(Boolean);

      let worst = 'green';
      for (const email of emails) {
        const st = await getStatusCached(email);
        worst = (st === 'red') ? 'red' : (st === 'yellow' && worst !== 'red') ? 'yellow' : worst;
        await sleep(THROTTLE_MS);
      }
      applyStyle(a, worst);
    }
  }

  function findPlainTextEmails(root = document) {
    if (!isEnabled()) return;
    if (getVisibleOnly() && document.hidden) return;

    const walker = document.createTreeWalker(root.body || root, NodeFilter.SHOW_TEXT, null);
    const nodes = [];
    let n; while ((n = walker.nextNode())) nodes.push(n);
    for (const node of nodes) {
      linkEmailsInTextNode(node);
    }
  }

  // ===== Scan-Planung =====
  function scheduleScan(delay = 150) {
    clearTimeout(window._pd_scan_t);
    window._pd_scan_t = setTimeout(() => { scan(); }, delay);
  }

  let scanning = false;
  async function scan() {
    if (scanning || !isEnabled()) return;
    if (getVisibleOnly() && document.hidden) return;
    scanning = true;
    try { findPlainTextEmails(document); await processMailtos(document); }
    finally { scanning = false; }
  }

  // ===== Wake/Resync & Sichtbarkeitswechsel =====
  function resync(reason) {
    try {
      startObserver();
      scheduleScan(50);
      console.log('[Pipedrive Highlighter] Resync:', reason);
    } catch (e) { console.warn('[Pipedrive Highlighter] Resync-Fehler:', e); }
  }
  window.addEventListener('focus', () => resync('focus'));
  window.addEventListener('online', () => resync('online'));
  document.addEventListener('visibilitychange', () => {
    if (!isEnabled()) return;
    if (getVisibleOnly()) {
      if (document.hidden) { stopObserver(); }
      else { startObserver(); scheduleScan(50); }
    }
  });

  // > 60s „Zeitloch“ ⇒ Wake (z.B. nach Ruhemodus)
  let lastTick = Date.now();
  setInterval(() => {
    const now = Date.now();
    if (now - lastTick > 60000) resync('wake');
    lastTick = now;
  }, 20000);

  // Schnell-Scan nach UI-Events (Overlays/Popups)
  ['click','hashchange','popstate'].forEach(ev =>
    window.addEventListener(ev, () => scheduleScan(120), true)
  );
  // Hotkey: ALT + S → Scan
  window.addEventListener('keydown', (e) => {
    if (e.altKey && (e.key || '').toLowerCase() === 's') {
      e.stopPropagation();
      scheduleScan(10);
    }
  }, true);

  // ===== Run =====
  async function run() {
    try {
      getToken();
      if (isEnabled()) { await scan(); startObserver(); console.log('[Pipedrive Highlighter] aktiv'); }
      else { stopObserver(); console.log('[Pipedrive Highlighter] AUS'); }
    } catch (e) { console.warn('[Pipedrive Highlighter]', e); }
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', run, { once: true });
  else run();

  // Debug-Helfer:
  window._pd_clearCaches = () => { cache.clear(); domainCache.clear(); persistNow(); console.log('[Pipedrive] Caches geleert'); };
})();
