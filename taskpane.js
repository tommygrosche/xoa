/*
 * xoa - xOrg Outlook Acceleration
 * Modul 1: Verwandte Mails (gleiche Conversation-ID oder exakt gleicher Betreff ohne Präfixe)
 */
(function () {
  'use strict';

  var VERSION = '1.10.1';
  var SETTINGS_KEY = 'xoaConfig';
  var Core = window.XoaCore;

  var DEFAULTS = {
    xobMailUrl: '',
    xobKey: '',
    maxResults: 25,
    minSubjectLength: 5,
    fetchTop: 50,
    searchMode: 'subject',
    prefixes: Core.DEFAULT_PREFIXES.slice()
  };

  var config = clone(DEFAULTS);
  var runToken = 0;
  var lastTopic = null;  // Suchparameter der aktuellen Mail für ein Themen-Fenster
  var lastSelection = '';
  var draftPrefixes = null;

  var $ = function (id) { return document.getElementById(id); };

  function clone(o) { return JSON.parse(JSON.stringify(o)); }

  // ---------- Konfiguration (roamingSettings: wandert mit dem Postfach, auch auf Android) ----------
  function loadConfig() {
    var stored = null;
    try { stored = Office.context.roamingSettings.get(SETTINGS_KEY); } catch (e) { stored = null; }
    config = clone(DEFAULTS);
    if (stored && typeof stored === 'object') {
      Object.keys(DEFAULTS).forEach(function (k) {
        if (stored[k] !== undefined && stored[k] !== null) config[k] = stored[k];
      });
    }
    config.prefixes = Core.cleanPrefixList(config.prefixes);
  }

  function saveConfig() {
    return new Promise(function (resolve, reject) {
      try {
        Office.context.roamingSettings.set(SETTINGS_KEY, config);
        Office.context.roamingSettings.saveAsync(function (r) {
          if (r.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(new Error((r.error && r.error.message) || 'Speichern fehlgeschlagen'));
        });
      } catch (e) { reject(e); }
    });
  }

  // ---------- xob-Aufruf ----------
  function uuid() {
    if (window.crypto && window.crypto.randomUUID) return window.crypto.randomUUID();
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
      var r = Math.random() * 16 | 0;
      return (c === 'x' ? r : (r & 0x3 | 0x8)).toString(16);
    });
  }

  function XobError(code, message, retryable) {
    this.code = code; this.message = message; this.retryable = !!retryable;
  }

  function fetchWithTimeout(url, options, ms) {
    if (!window.AbortController) return fetch(url, options);
    var ctrl = new AbortController();
    var timer = setTimeout(function () { ctrl.abort(); }, ms);
    options.signal = ctrl.signal;
    return fetch(url, options).then(function (r) { clearTimeout(timer); return r; },
      function (e) { clearTimeout(timer); throw e; });
  }

  function callXobOnce(service, payload) {
    var requestId = uuid();
    var body = { service: service, version: 1, requestId: requestId, client: 'xoa', key: config.xobKey, payload: payload || {} };
    var started = Date.now();
    return fetchWithTimeout(config.xobMailUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    }, 60000).then(function (r) {
      return r.text().then(function (text) {
        var json = null;
        try { json = text ? JSON.parse(text) : null; } catch (e) { json = null; }
        var ms = Date.now() - started;
        if (r.status === 401) throw new XobError('UNAUTHORIZED', 'Der xob-Schlüssel stimmt nicht. Prüf ihn in den Einstellungen.', false);
        if (!r.ok || !json || json.status === 'error') {
          var msg = (json && json.error && json.error.message) || ('xob antwortet mit HTTP ' + r.status + '.');
          throw new XobError((json && json.error && json.error.code) || ('HTTP_' + r.status), msg, r.status === 429 || r.status === 503 || r.status === 504);
        }
        json._ms = ms;
        return json;
      });
    }, function (e) {
      var aborted = e && e.name === 'AbortError';
      throw new XobError(aborted ? 'TIMEOUT' : 'NETWORK',
        aborted ? 'xob hat nicht rechtzeitig geantwortet.' : 'xob ist nicht erreichbar. Netzwerk oder URL prüfen.', true);
    });
  }

  function callXob(service, payload) {
    return callXobOnce(service, payload).catch(function (e) {
      if (!e.retryable) throw e;
      return new Promise(function (res) { setTimeout(res, 800); }).then(function () { return callXobOnce(service, payload); });
    });
  }

  // ---------- Aktuelle Mail ----------
  function currentItem() {
    var mb = Office.context.mailbox;
    return mb ? mb.item : null;
  }

  function currentRestId(item) {
    try {
      var mb = Office.context.mailbox;
      if (item && item.itemId && mb.convertToRestId) return mb.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0);
    } catch (e) { /* Mobile liefert ggf. bereits REST-IDs */ }
    return item ? item.itemId : null;
  }

  // Ausgewählte Mail als Graph-ähnliches Objekt (Ersatz, falls xob sie nicht mitliefert)
  function currentAsMessage(item, restId) {
    function addr(a) { return a ? { emailAddress: { name: a.displayName || a.emailAddress || '', address: a.emailAddress || '' } } : null; }
    var created = item.dateTimeCreated instanceof Date ? item.dateTimeCreated.toISOString() : (item.dateTimeCreated || '');
    return {
      id: restId || item.itemId || 'current',
      internetMessageId: item.internetMessageId || '',
      subject: item.subject || '',
      from: addr(item.from),
      toRecipients: (item.to || []).map(addr).filter(Boolean),
      receivedDateTime: created,
      hasAttachments: !!(item.attachments && item.attachments.length),
      flag: { flagStatus: 'notFlagged' },
      webLink: restId ? 'https://outlook.office.com/mail/deeplink/read/' + encodeURIComponent(restId) : ''
    };
  }

  // ---------- Darstellung ----------
  var FLAG_SVG = '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M3.5 14.5V2.5m0 0h8l-1.8 3 1.8 3h-8" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>';
  var FLAG_DONE_SVG = '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M3 8.5l3 3 7-7" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>';

  function setState(text, kind, actionLabel, action) {
    var el = $('state');
    el.className = 'state' + (kind ? ' ' + kind : '');
    el.textContent = text || '';
    if (actionLabel && action) {
      var b = document.createElement('button');
      b.type = 'button'; b.className = 'btn btn-secondary btn-small'; b.textContent = actionLabel;
      b.addEventListener('click', action);
      el.appendChild(document.createElement('br'));
      el.appendChild(b);
    }
  }

  function showSkeleton() {
    var list = $('list');
    list.innerHTML = '';
    for (var i = 0; i < 4; i++) {
      var s = document.createElement('div'); s.className = 'skeleton'; list.appendChild(s);
    }
  }

  function span(cls, text) {
    var s = document.createElement('span'); s.className = cls; s.textContent = text; return s;
  }

  function renderList(result) {
    var list = $('list');
    list.innerHTML = '';
    result.items.forEach(function (entry) {
      var m = entry.message;
      var row = document.createElement('div');
      row.className = 'item' + (entry.current ? ' current' : '');
      row.setAttribute('role', 'listitem');
      row.tabIndex = 0;

      var l1 = document.createElement('div');
      l1.className = 'line1';
      var from = Core.personName(m.from) || '(unbekannt)';
      var to = Core.recipientsLabel(m.toRecipients);
      l1.appendChild(span('from', from));
      if (entry.current) l1.appendChild(span('cur-badge', 'Ausgewählt'));
      if (to) l1.appendChild(span('to', to));
      else l1.appendChild(span('to', ''));
      l1.appendChild(span('time', Core.formatTimestamp(m.receivedDateTime || m.sentDateTime)));
      var f = document.createElement('span');
      var st = Core.flagState(m);
      f.className = 'flag ' + st;
      if (st === 'flagged') { f.innerHTML = FLAG_SVG; f.title = 'Zur Nachverfolgung gekennzeichnet'; }
      else if (st === 'complete') { f.innerHTML = FLAG_DONE_SVG; f.title = 'Nachverfolgung erledigt'; }
      l1.appendChild(f);

      var l2 = document.createElement('div');
      l2.className = 'line2';
      l2.textContent = m.subject || '(ohne Betreff)';
      l2.title = m.subject || '';

      row.appendChild(l1);
      row.appendChild(l2);
      row.setAttribute('aria-label', from + ', ' + Core.formatTimestamp(m.receivedDateTime || m.sentDateTime) + ', ' + (m.subject || ''));

      row.addEventListener('click', function () {
        var sel = window.getSelection ? String(window.getSelection()) : '';
        if (sel.trim()) return; // Text wird markiert, nicht öffnen
        openMessage(m);
      });
      row.addEventListener('keydown', function (ev) {
        if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); openMessage(m); }
      });
      list.appendChild(row);
    });
  }

  function toast(text) {
    var t = $('toast');
    t.textContent = text;
    t.hidden = false;
    clearTimeout(toast._t);
    toast._t = setTimeout(function () { t.hidden = true; }, 3000);
  }

  // ---------- Themen-Fenster (eigenständige Browserseite, je Thema eines) ----------
  function canUseWindow() {
    return !isMobile();
  }

  function b64(obj) {
    return btoa(unescape(encodeURIComponent(JSON.stringify(obj))));
  }

  function viewerUrl(topic, offset, how) {
    var base = location.href.split('#')[0].split('?')[0].replace(/[^/]*$/, 'viewer.html');
    var cfg = { u: config.xobMailUrl, k: config.xobKey };
    return base + '#o=' + (offset || 0) + '&m=' + (how || 'window.open') + '&t=' + encodeURIComponent(b64(topic)) + '&c=' + encodeURIComponent(b64(cfg));
  }

  // Größe und Lage des letzten Fensters (vom Fenster selbst im localStorage abgelegt) und
  // Zahl der schon offenen Fenster: das nächste Fenster öffnet an derselben Stelle, je offenem
  // Fenster 30 Pixel nach rechts unten versetzt.
  var WINDOW_STEP = 30;
  // Startgröße ohne gemerkte Größe: 1650 x 930, höchstens 92 % des Bildschirms
  function startW() { return Math.round(Math.min(1650, ((window.screen && window.screen.availWidth) || 1920) * 0.92)); }
  function startH() { return Math.round(Math.min(930, ((window.screen && window.screen.availHeight) || 1080) * 0.92)); }
  function windowFeatures() {
    var g = null, open = {}, off = 0;
    try { g = JSON.parse(localStorage.getItem('xoaViewerGeom') || 'null'); } catch (e) { g = null; }
    try { open = JSON.parse(localStorage.getItem('xoaViewerOpen') || '{}') || {}; } catch (e) { open = {}; }
    var now = Date.now();
    var alive = Object.keys(open).filter(function (k) { return now - open[k] < 6000; }).length;
    if (!g || !(g.w > 300 && g.h > 200)) return { offset: 0, features: 'popup=yes,width=' + startW() + ',height=' + startH() };
    off = alive * WINDOW_STEP;
    var maxX = Math.max(0, (window.screen.availWidth || 1920) - g.w);
    var maxY = Math.max(0, (window.screen.availHeight || 1080) - g.h);
    var x = Math.round(g.x) + off, y = Math.round(g.y) + off;
    if (x > maxX || y > maxY) { x = Math.round(g.x); y = Math.round(g.y); off = 0; }
    x = Math.max(0, Math.min(x, maxX)); y = Math.max(0, Math.min(y, maxY));
    return { offset: off, features: 'popup=yes,width=' + Math.round(g.w) + ',height=' + Math.round(g.h) + ',left=' + x + ',top=' + y };
  }

  // Die frühere Auswahl "Fenster / Browser-Tab" ist entfernt (Outlook entscheidet selbst, wo das Fenster aufgeht).
  // Auswahl und Hinweistext darunter werden ausgeblendet, falls die Seite sie noch enthält.
  function removeWindowControlUi() {
    try { var sel = $('windowControl'); if (sel && sel.parentNode) sel.parentNode.removeChild(sel); } catch (e) { /* egal */ }
    try {
      var btn = $('btnOpenWindow');
      var scope = (btn && btn.parentNode) || document.body;
      Array.prototype.forEach.call(scope.querySelectorAll('p,div,span,small'), function (el) {
        if (el.children.length === 0 && /^\s*Öffnet ein eigenes Fenster/.test(el.textContent || '')) el.style.display = 'none';
      });
    } catch (e) { /* egal */ }
  }

  function openWindow() {
    if (!lastTopic) { toast('Wähle zuerst eine Mail aus.'); return; }
    var w = windowFeatures();
    var url = viewerUrl(lastTopic, w.offset);
    // Eigenes Fenster (merkt sich Größe und Lage); wird es blockiert, bleibt der Browser-Tab als Rückfall
    var win = null;
    try { win = window.open(url, 'xoaRelated' + Date.now(), w.features); } catch (e) { win = null; }
    if (win) return;
    try {
      if (Office.context.ui && Office.context.ui.openBrowserWindow) { Office.context.ui.openBrowserWindow(viewerUrl(lastTopic, 0, 'openBrowserWindow')); return; }
    } catch (e) { /* weiter */ }
    toast('Das Fenster wurde blockiert. Erlaube Popups für dieses Add-in.');
  }

  // ---------- Mail öffnen ----------
  function isMobile() {
    try {
      var p = Office.context.diagnostics && Office.context.diagnostics.platform;
      return p === Office.PlatformType.Android || p === Office.PlatformType.iOS || p === 'Android' || p === 'iOS';
    } catch (e) { return false; }
  }

  function openWeb(url) {
    if (!url) { toast('Die Mail lässt sich hier nicht öffnen.'); return; }
    try {
      if (Office.context.ui && Office.context.ui.openBrowserWindow) { Office.context.ui.openBrowserWindow(url); return; }
    } catch (e) { /* weiter mit window.open */ }
    window.open(url, '_blank', 'noopener');
  }

  function openMessage(m) {
    var mb = Office.context.mailbox;
    var id = null;
    if (isMobile()) {
      id = m.id;
    } else {
      try { id = mb.convertToEwsId(m.id, Office.MailboxEnums.RestVersion.v2_0); } catch (e) { id = null; }
    }
    if (!id) { openWeb(m.webLink); return; }
    try {
      if (mb.displayMessageFormAsync) {
        mb.displayMessageFormAsync(id, function (r) {
          if (r.status !== Office.AsyncResultStatus.Succeeded) openWeb(m.webLink);
        });
      } else if (mb.displayMessageForm) {
        mb.displayMessageForm(id);
      } else {
        openWeb(m.webLink);
      }
    } catch (e) { openWeb(m.webLink); }
  }

  // ---------- Schalter "Volltext" (Suchtext nur im Betreff oder auch im Mailtext) ----------
  var FULL_KEY = 'xoaFullText';
  function isFullText() {
    var box = $('fullText');
    return !!(box && box.checked);
  }
  function buildFullTextSwitch() {
    var field = $('terms');
    if (!field || $('fullText')) return;
    var saved = null;
    try { saved = localStorage.getItem(FULL_KEY); } catch (e) { saved = null; }
    var style = document.createElement('style');
    style.textContent =
      '.xoa-switch{display:inline-flex;align-items:center;gap:6px;margin:4px 0;font-size:12px;cursor:pointer;user-select:none}' +
      '.xoa-switch input{position:absolute;opacity:0;width:0;height:0}' +
      '.xoa-switch .tr{position:relative;width:30px;height:16px;border-radius:8px;background:#b9bec7;flex:0 0 auto}' +
      '.xoa-switch .tr::after{content:"";position:absolute;top:2px;left:2px;width:12px;height:12px;border-radius:50%;background:#fff}' +
      '.xoa-switch input:checked+.tr{background:#0f6cbd}' +
      '.xoa-switch input:checked+.tr::after{left:16px}' +
      '.xoa-switch input:focus-visible+.tr{outline:2px solid #0f6cbd;outline-offset:2px}';
    document.head.appendChild(style);
    var label = document.createElement('label');
    label.className = 'xoa-switch';
    label.title = 'Aus: Der Suchtext wird nur im Betreff gesucht. An: Volltextsuche, auch im Text der Mail (langsamer, mehr Treffer).';
    var box = document.createElement('input');
    box.type = 'checkbox';
    box.id = 'fullText';
    box.checked = saved === null ? config.searchMode === 'all' : saved === '1';
    var track = document.createElement('span');
    track.className = 'tr';
    var text = document.createElement('span');
    text.textContent = 'Volltext';
    label.appendChild(box); label.appendChild(track); label.appendChild(text);
    field.insertAdjacentElement('afterend', label);
    box.addEventListener('change', function () {
      try { localStorage.setItem(FULL_KEY, box.checked ? '1' : '0'); } catch (e) { /* egal */ }
      refresh();
    });
  }

  // ---------- Hauptablauf ----------
  function refresh() {
    var token = ++runToken;
    var item = currentItem();
    $('list').innerHTML = '';
    $('resultCount').textContent = '';

    $('windowBox').hidden = true;
    lastTopic = null;
    if (!item) {
      $('normSubject').textContent = '–';
      $('rawSubject').textContent = '–';
      setState('Wähle eine Mail aus.');
      return;
    }

    var raw = item.subject || '';
    var normalized = Core.normalizeSubject(raw, config.prefixes);
    $('rawSubject').textContent = raw || '(ohne Betreff)';
    $('normSubject').textContent = normalized || '(ohne Betreff)';

    if (!config.xobMailUrl || !config.xobKey) {
      setState('xob ist noch nicht eingerichtet.', null, 'Einstellungen öffnen', openSettings);
      return;
    }

    var useSubject = normalized.length >= (parseInt(config.minSubjectLength, 10) || 5);
    var conversationId = item.conversationId || '';
    var terms = Core.parseTerms($('terms').value);
    var fullText = isFullText();
    if (!conversationId && !useSubject && !terms.length) {
      setState('Betreff zu kurz und keine Conversation-ID vorhanden.');
      return;
    }

    lastTopic = {
      v: 1,
      conversationId: conversationId,
      normalized: normalized,
      raw: raw,
      useSubject: useSubject,
      mode: config.searchMode === 'all' ? 'all' : 'subject',
      top: Math.min(100, Math.max(parseInt(config.fetchTop, 10) || 50, (parseInt(config.maxResults, 10) || 25) * 2)),
      maxResults: config.maxResults,
      prefixes: config.prefixes,
      terms: terms,
      fullText: fullText,
      curImid: item.internetMessageId || '',
      curRest: currentRestId(item) || '',
      cur: currentAsMessage(item, currentRestId(item))
    };
    setState('Suche läuft…');
    if (canUseWindow()) $('list').innerHTML = ''; else showSkeleton();

    Core.relatedQuery(callXob, {
      conversationId: conversationId,
      normalized: normalized,
      useSubject: useSubject,
      mode: lastTopic.mode,
      terms: terms,
      fullText: fullText,
      top: lastTopic.top
    }).then(function (json) {
      if (token !== runToken) return;
      var data = json.data || {};
      var result = Core.mergeResults(data.conversation || [], data.subject || [], {
        normalized: normalized,
        prefixes: config.prefixes,
        currentInternetMessageId: item.internetMessageId,
        currentRestId: currentRestId(item),
        currentMessage: lastTopic.cur,
        maxResults: config.maxResults,
        useSubject: useSubject,
        terms: terms,
        fullText: fullText,
        termHits: data.terms || []
      });
      $('meta').textContent = 'xob · ' + json._ms + ' ms' + (useSubject ? '' : ' · Betreff zu kurz, nur Conversation') +
        (terms.length ? ' · ' + (fullText ? 'Volltext' : 'Betreff') + ': ' + terms.join('; ') : '');
      var countLabel = result.total > result.items.length
        ? result.items.length + ' von ' + result.total
        : String(result.items.length);
      if (!result.items.length) {
        $('list').innerHTML = '';
        setState('Keine verwandten Mails gefunden.');
        return;
      }
      $('resultCount').textContent = countLabel;
      if (canUseWindow()) {
        $('list').innerHTML = '';
        setState(result.total === 1 ? '1 Mail gefunden (nur die ausgewählte).' : result.total + ' Mails gefunden, inklusive der ausgewählten.');
        $('windowBox').hidden = false;
      } else {
        setState('');
        renderList(result);
      }
    }).catch(function (e) {
      if (token !== runToken) return;
      $('list').innerHTML = '';
      setState(e.message || 'Unbekannter Fehler.', 'error', 'Erneut versuchen', refresh);
    });
  }

  // ---------- Präfix aus Markierung ----------
  function trackSelection() {
    var sel = window.getSelection ? String(window.getSelection()).replace(/\s+/g, ' ').trim() : '';
    if (sel && sel.length <= 40) {
      lastSelection = sel;
      $('selectionHint').textContent = 'Markiert: „' + sel + '“';
    }
  }

  function addPrefixFromSelection() {
    if (!lastSelection) {
      $('selectionHint').textContent = 'Markiere zuerst Text im Originalbetreff oder in einem Eintrag.';
      return;
    }
    addPrefix(lastSelection, true);
  }

  function addPrefix(p, persist) {
    var t = String(p || '').replace(/\s+/g, ' ').trim();
    if (!t) return false;
    var exists = config.prefixes.some(function (x) { return x.toLowerCase() === t.toLowerCase(); });
    if (exists) { toast('„' + t + '“ ist schon hinterlegt.'); return false; }
    config.prefixes.push(t);
    config.prefixes = Core.cleanPrefixList(config.prefixes);
    if (persist) {
      saveConfig().then(function () {
        toast('Präfix „' + t + '“ hinzugefügt');
        lastSelection = '';
        $('selectionHint').textContent = '';
        refresh();
      }, function (e) { toast('Speichern fehlgeschlagen: ' + e.message); });
    }
    return true;
  }

  // ---------- Einstellungen ----------
  function renderPrefixChips() {
    var box = $('prefixList');
    box.innerHTML = '';
    draftPrefixes.forEach(function (p, idx) {
      var c = document.createElement('span');
      c.className = 'chip';
      c.appendChild(document.createTextNode(p));
      var x = document.createElement('button');
      x.type = 'button';
      x.textContent = '×';
      x.setAttribute('aria-label', 'Präfix ' + p + ' entfernen');
      x.addEventListener('click', function () { draftPrefixes.splice(idx, 1); renderPrefixChips(); });
      c.appendChild(x);
      box.appendChild(c);
    });
  }

  function openSettings() {
    $('cfgUrl').value = config.xobMailUrl || '';
    $('cfgKey').value = config.xobKey || '';
    $('cfgMax').value = config.maxResults;
    $('cfgMin').value = config.minSubjectLength;
    $('cfgMode').value = config.searchMode === 'all' ? 'all' : 'subject';
    draftPrefixes = config.prefixes.slice();
    renderPrefixChips();
    ['errUrl', 'errKey', 'errMax', 'errMin', 'errPrefix'].forEach(function (id) { $(id).textContent = ''; });
    $('settingsStatus').textContent = '';
    $('viewMain').hidden = true;
    $('viewSettings').hidden = false;
  }

  function closeSettings() {
    $('viewSettings').hidden = true;
    $('viewMain').hidden = false;
  }

  function readIntField(id, errId, min, max, label) {
    var v = $(id).value.trim();
    var n = Number(v);
    if (!v || !Number.isInteger(n) || n < min || n > max) {
      $(errId).textContent = label + ' muss eine ganze Zahl zwischen ' + min + ' und ' + max + ' sein.';
      return null;
    }
    $(errId).textContent = '';
    return n;
  }

  function collectSettings() {
    var ok = true;
    var url = $('cfgUrl').value.trim();
    if (url && !/^https:\/\//i.test(url)) { $('errUrl').textContent = 'Die URL muss mit https:// beginnen.'; ok = false; }
    else $('errUrl').textContent = '';
    var key = $('cfgKey').value.trim();
    if (url && !key) { $('errKey').textContent = 'Trag den xob-Schlüssel ein.'; ok = false; }
    else $('errKey').textContent = '';
    var max = readIntField('cfgMax', 'errMax', 1, 100, 'Max. Treffer');
    var min = readIntField('cfgMin', 'errMin', 1, 50, 'Min. Zeichen');
    if (max === null || min === null) ok = false;
    if (!ok) return null;
    return { xobMailUrl: url, xobKey: key, maxResults: max, minSubjectLength: min, searchMode: $('cfgMode').value, prefixes: Core.cleanPrefixList(draftPrefixes) };
  }

  function saveSettings() {
    var s = collectSettings();
    if (!s) return;
    Object.keys(s).forEach(function (k) { config[k] = s[k]; });
    $('settingsStatus').textContent = 'Speichert…';
    saveConfig().then(function () {
      $('settingsStatus').textContent = '';
      toast('Einstellungen gespeichert');
      closeSettings();
      refresh();
    }, function (e) { $('settingsStatus').textContent = 'Speichern fehlgeschlagen: ' + e.message; });
  }

  function testConnection() {
    var s = collectSettings();
    if (!s) return;
    if (!s.xobMailUrl) { $('errUrl').textContent = 'Trag zuerst die URL ein.'; return; }
    var saved = { url: config.xobMailUrl, key: config.xobKey };
    config.xobMailUrl = s.xobMailUrl; config.xobKey = s.xobKey;
    $('settingsStatus').textContent = 'Teste Verbindung…';
    callXobOnce('system.ping', {}).then(function (j) {
      var v = j.data && j.data.version ? ' · ' + j.data.version : '';
      $('settingsStatus').textContent = 'Verbunden in ' + j._ms + ' ms' + v;
    }, function (e) {
      $('settingsStatus').textContent = e.message;
    }).then(function () { config.xobMailUrl = saved.url; config.xobKey = saved.key; });
  }

  function bindUi() {
    buildFullTextSwitch();
    removeWindowControlUi();
    $('btnRefresh').addEventListener('click', refresh);
    $('btnOpenWindow').addEventListener('click', openWindow);
    $('termForm').addEventListener('submit', function (ev) { ev.preventDefault(); refresh(); });
    $('btnSettings').addEventListener('click', function () { if ($('viewSettings').hidden) openSettings(); else closeSettings(); });
    $('btnClose').addEventListener('click', closeSettings);
    $('btnSave').addEventListener('click', saveSettings);
    $('btnTest').addEventListener('click', testConnection);
    $('btnShowKey').addEventListener('click', function () {
      var k = $('cfgKey');
      var show = k.type === 'password';
      k.type = show ? 'text' : 'password';
      $('btnShowKey').textContent = show ? 'Verbergen' : 'Anzeigen';
    });
    $('btnAddPrefix').addEventListener('click', function () {
      var v = $('cfgNewPrefix').value.replace(/\s+/g, ' ').trim();
      if (!v) { $('errPrefix').textContent = 'Gib einen Präfix ein.'; return; }
      if (draftPrefixes.some(function (x) { return x.toLowerCase() === v.toLowerCase(); })) {
        $('errPrefix').textContent = 'Dieser Präfix ist schon hinterlegt.'; return;
      }
      $('errPrefix').textContent = '';
      draftPrefixes.push(v);
      $('cfgNewPrefix').value = '';
      renderPrefixChips();
    });
    $('cfgNewPrefix').addEventListener('input', function () { $('errPrefix').textContent = ''; });
    $('btnAddSelection').addEventListener('click', addPrefixFromSelection);
    document.addEventListener('selectionchange', trackSelection);
    $('versionInfo').textContent = 'xoa ' + VERSION;
    $('appVersion').textContent = 'v' + VERSION;
  }

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Outlook) {
      setState('xoa läuft nur in Outlook.', 'error');
      return;
    }
    loadConfig();
    bindUi();
    try {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, function () { closeSettings(); $('terms').value = ''; refresh(); });
    } catch (e) { /* ohne Pinning */ }
    refresh();
  });
})();
