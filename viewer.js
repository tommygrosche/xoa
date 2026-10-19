/*
 * xoa - Themen-Fenster "Verwandte Mails"
 * Eigenständige Browserseite: je Thema ein Fenster/Tab, unabhängig von Outlook.
 * Bekommt Suchparameter (#t=) und xob-Zugang (#c=) vom Panel, entfernt beides sofort aus der Adresse
 * und fragt xob selbst ab. Doppelklick öffnet die Mail in Outlook im Web.
 */
(function () {
  'use strict';

  var Core = window.XoaCore;
  var $ = function (id) { return document.getElementById(id); };
  var CFG_KEY = 'xoaViewerCfg';
  var TOPIC_PREFIX = 'xoaTopic:';

  var topic = null;
  var cfg = null;
  var view = { kind: 'loading' };
  var sortKey = 'received';
  var sortDir = -1;
  var selectedId = null;

  var CLIP = '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M10.5 4.5 5.8 9.2a1.4 1.4 0 0 0 2 2l5-5a2.8 2.8 0 0 0-4-4l-5.2 5.2a4.2 4.2 0 0 0 6 6l4.2-4.2" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"/></svg>';
  var FLAG = '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M3.5 14.5V2.5m0 0h8l-1.8 3 1.8 3h-8" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>';
  var DONE = '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M3 8.5l3 3 7-7" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>';

  // ---------- Parameter, Speicher ----------
  function unb64(s) {
    try { return JSON.parse(decodeURIComponent(escape(atob(s)))); } catch (e) { return null; }
  }
  function b64(obj) { return btoa(unescape(encodeURIComponent(JSON.stringify(obj)))); }

  function store(kind, key, value) {
    try { (kind === 'local' ? localStorage : sessionStorage).setItem(key, JSON.stringify(value)); } catch (e) { /* Speicher gesperrt */ }
  }
  function load(kind, key) {
    try { var v = (kind === 'local' ? localStorage : sessionStorage).getItem(key); return v ? JSON.parse(v) : null; } catch (e) { return null; }
  }

  function readParams() {
    var hash = location.hash.replace(/^#/, '');
    var params = {};
    hash.split('&').forEach(function (part) {
      var i = part.indexOf('=');
      if (i > 0) params[part.slice(0, i)] = decodeURIComponent(part.slice(i + 1));
    });
    if (params.c) {
      var c = unb64(params.c);
      if (c && c.u && c.k) { cfg = c; store('local', CFG_KEY, c); }
    }
    if (!cfg) cfg = load('local', CFG_KEY);
    if (params.t) topic = unb64(params.t);
    // Zugangsdaten nie in Adresse/Verlauf stehen lassen; Thema bleibt für Neuladen erhalten
    if (topic) {
      var id = (params.id) || String(Date.now()) + Math.random().toString(16).slice(2, 6);
      store('session', TOPIC_PREFIX + id, topic);
      try { history.replaceState(null, '', location.pathname + '#id=' + id + '&t=' + encodeURIComponent(b64(topic))); } catch (e) { /* egal */ }
    } else if (params.id) {
      topic = load('session', TOPIC_PREFIX + params.id);
    }
  }

  // ---------- xob ----------
  function callXob(service, payload) {
    var started = Date.now();
    var body = {
      service: service, version: 1, client: 'xoa-viewer', key: cfg.k, payload: payload,
      requestId: (window.crypto && crypto.randomUUID) ? crypto.randomUUID() : String(started)
    };
    return fetch(cfg.u, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) })
      .then(function (r) {
        return r.text().then(function (text) {
          var json = null;
          try { json = text ? JSON.parse(text) : null; } catch (e) { json = null; }
          if (r.status === 401) throw { code: 'UNAUTHORIZED', message: 'Der xob-Schlüssel stimmt nicht. Prüf ihn in den Einstellungen des Panels.' };
          if (!r.ok || !json || json.status === 'error') {
            throw { code: (json && json.error && json.error.code) || ('HTTP_' + r.status), message: (json && json.error && json.error.message) || ('xob antwortet mit HTTP ' + r.status + '.') };
          }
          json._ms = Date.now() - started;
          return json;
        });
      }, function () {
        throw { code: 'NETWORK', message: 'xob ist nicht erreichbar. Netzwerk prüfen und erneut versuchen.' };
      });
  }

  function search() {
    view = { kind: 'loading' };
    render();
    function run(mode) {
      return callXob('mail.related', {
        conversationId: topic.conversationId || '',
        search: topic.useSubject ? Core.buildSearch(topic.normalized, mode) : '',
        top: topic.top || 50
      });
    }
    run(topic.mode || 'subject').catch(function (e) {
      if (topic.useSubject && (topic.mode || 'subject') === 'subject' && e.code !== 'UNAUTHORIZED' && e.code !== 'NETWORK') return run('all');
      throw e;
    }).then(function (json) {
      var data = json.data || {};
      var result = Core.mergeResults(data.conversation || [], data.subject || [], {
        normalized: topic.normalized,
        prefixes: topic.prefixes || Core.DEFAULT_PREFIXES,
        currentInternetMessageId: topic.curImid,
        currentRestId: topic.curRest,
        currentMessage: topic.cur || null,
        maxResults: topic.maxResults || 25,
        useSubject: topic.useSubject
      });
      view = {
        kind: 'results', ms: json._ms, conversationOnly: !topic.useSubject,
        count: result.total > result.items.length ? result.items.length + ' von ' + result.total : String(result.items.length),
        total: result.total,
        items: result.items.map(function (e) { var m = Object.assign({}, e.message); m.__current = !!e.current; return m; })
      };
      render();
    }).catch(function (e) {
      view = { kind: 'error', text: (e && e.message) || 'Unbekannter Fehler.' };
      render();
    });
  }

  // ---------- Darstellung ----------
  function setState(text, isError) {
    var s = $('state');
    s.textContent = text || '';
    s.className = 'state' + (isError ? ' error' : '');
  }

  function fields(m) {
    return {
      id: m.id,
      webLink: m.webLink || '',
      from: Core.personName(m.from) || '',
      fromAddress: (m.from && m.from.emailAddress && m.from.emailAddress.address) || '',
      to: (m.toRecipients || []).map(Core.personName).filter(Boolean).join('; '),
      subject: m.subject || '',
      received: m.receivedDateTime || m.sentDateTime || '',
      hasAttachments: !!m.hasAttachments,
      flag: Core.flagState(m),
      current: !!m.__current
    };
  }

  function sortValue(it, key) {
    if (key === 'received') { var t = Date.parse(it.received); return isNaN(t) ? 0 : t; }
    if (key === 'hasAttachments') return it.hasAttachments ? 1 : 0;
    if (key === 'flag') return it.flag === 'flagged' ? 2 : (it.flag === 'complete' ? 1 : 0);
    return String(it[key] || '').toLowerCase();
  }

  function visibleItems() {
    var items = (view.items || []).map(fields);
    var q = $('filter').value.trim().toLowerCase();
    if (q) {
      items = items.filter(function (it) {
        return (it.from + ' ' + it.fromAddress + ' ' + it.to + ' ' + it.subject).toLowerCase().indexOf(q) >= 0;
      });
    }
    items.sort(function (a, b) {
      var x = sortValue(a, sortKey), y = sortValue(b, sortKey);
      if (x < y) return -1 * sortDir;
      if (x > y) return 1 * sortDir;
      return sortValue(b, 'received') - sortValue(a, 'received');
    });
    return items;
  }

  function td(cls, text, title) {
    var c = document.createElement('td');
    c.className = cls;
    c.textContent = text;
    if (title) c.title = title;
    return c;
  }

  function openItem(it) {
    if (!it.webLink) { setState('Für diese Mail liegt kein Link vor.', true); return; }
    window.open(it.webLink, '_blank', 'noopener');
  }

  function select(tr, it) {
    var prev = document.querySelector('tbody tr.sel');
    if (prev && prev !== tr) { prev.classList.remove('sel'); prev.setAttribute('aria-selected', 'false'); }
    tr.classList.add('sel');
    tr.setAttribute('aria-selected', 'true');
    selectedId = it.id;
  }

  function render() {
    var subj = (topic && topic.normalized) || '–';
    $('subject').textContent = subj;
    $('subject').title = (topic && topic.raw) || subj;
    document.title = 'xOrg – ' + subj;
    var body = $('rows');
    body.innerHTML = '';
    $('count').textContent = '';
    $('meta').textContent = '';

    if (view.kind === 'loading') { setState('Suche läuft…'); return; }
    if (view.kind === 'message') { setState(view.text || ''); return; }
    if (view.kind === 'error') { setState(view.text || 'Unbekannter Fehler.', true); return; }

    var items = visibleItems();
    var total = (view.items || []).length;
    $('count').textContent = items.length === total ? (view.count || String(total)) : items.length + ' von ' + total;
    $('meta').textContent = 'xob · ' + view.ms + ' ms' + (view.conversationOnly ? ' · nur Conversation (Betreff zu kurz)' : '');

    if (!total) { setState('Keine Mails gefunden.'); return; }
    if (!items.length) { setState('Kein Treffer für diesen Filter.'); return; }
    setState('');

    items.forEach(function (it) {
      var tr = document.createElement('tr');
      tr.tabIndex = 0;
      tr.setAttribute('aria-selected', 'false');
      var fromCell = td('from', '', it.fromAddress || it.from);
      if (it.current) {
        tr.classList.add('current');
        var badge = document.createElement('span');
        badge.className = 'cur-badge';
        badge.textContent = 'Ausgewählt';
        fromCell.appendChild(badge);
      }
      fromCell.appendChild(document.createTextNode(it.from || '(unbekannt)'));
      tr.appendChild(fromCell);
      tr.appendChild(td('to', it.to, it.to));
      tr.appendChild(td('subject', it.subject || '(ohne Betreff)', it.subject));
      tr.appendChild(td('received', Core.formatTimestamp(it.received)));
      var a = td('ic att', '');
      if (it.hasAttachments) { a.innerHTML = CLIP; a.title = 'Mit Anhang'; }
      tr.appendChild(a);
      var f = td('ic flag ' + it.flag, '');
      if (it.flag === 'flagged') { f.innerHTML = FLAG; f.title = 'Zur Nachverfolgung gekennzeichnet'; }
      else if (it.flag === 'complete') { f.innerHTML = DONE; f.title = 'Erledigt'; }
      tr.appendChild(f);

      tr.addEventListener('click', function () { select(tr, it); });
      tr.addEventListener('dblclick', function () { select(tr, it); openItem(it); });
      tr.addEventListener('keydown', function (ev) {
        if (ev.key === 'Enter') { ev.preventDefault(); openItem(it); }
        else if (ev.key === 'ArrowDown' && tr.nextSibling) { ev.preventDefault(); tr.nextSibling.focus(); }
        else if (ev.key === 'ArrowUp' && tr.previousSibling) { ev.preventDefault(); tr.previousSibling.focus(); }
      });
      tr.addEventListener('focus', function () { select(tr, it); });
      if (it.id === selectedId) { tr.classList.add('sel'); tr.setAttribute('aria-selected', 'true'); }
      body.appendChild(tr);
    });
  }

  function bindHeader() {
    Array.prototype.forEach.call(document.querySelectorAll('thead th'), function (th) {
      th.querySelector('button').addEventListener('click', function () {
        var key = th.getAttribute('data-key');
        if (sortKey === key) sortDir = -sortDir;
        else { sortKey = key; sortDir = (key === 'received' || key === 'hasAttachments' || key === 'flag') ? -1 : 1; }
        Array.prototype.forEach.call(document.querySelectorAll('thead th'), function (o) { o.removeAttribute('aria-sort'); });
        th.setAttribute('aria-sort', sortDir === 1 ? 'ascending' : 'descending');
        render();
      });
    });
  }

  function start() {
    bindHeader();
    $('filter').addEventListener('input', render);
    $('btnRefresh').addEventListener('click', function () { if (topic && cfg) search(); });
    readParams();
    var now = new Date();
    $('opened').textContent = 'geöffnet ' + Core.formatTimestamp(now.toISOString());
    if (!topic) { view = { kind: 'message', text: 'Kein Thema übergeben. Öffne dieses Fenster aus Outlook über xOrg → „In Fenster anzeigen“.' }; render(); return; }
    if (!cfg) { view = { kind: 'message', text: 'xob-Zugang fehlt. Öffne dieses Fenster erneut aus Outlook.' }; render(); return; }
    search();
  }

  start();
})();
