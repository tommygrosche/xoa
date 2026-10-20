/*
 * xoa - Anbindung des Fensters "Related emails" (viewer.js / viewer.css sind dieselben Dateien
 * wie in xOrg Desktop). Diese Datei stellt window.related bereit:
 * Suche gegen xob, Mail öffnen, Größe und Lage des Fensters merken.
 *
 * Das Panel übergibt Suchparameter (#t=) und xob-Zugang (#c=); beides wird sofort aus der
 * Adresse entfernt. Größe und Lage liegen im localStorage (gleicher Ursprung wie das Panel),
 * damit das Panel das nächste Fenster an derselben Stelle öffnen kann.
 */
(function () {
  'use strict';

  var Core = window.XoaCore;
  var CFG_KEY = 'xoaViewerCfg';
  var TOPIC_PREFIX = 'xoaTopic:';
  var GEOM_KEY = 'xoaViewerGeom';
  var OPEN_KEY = 'xoaViewerOpen';

  var topic = null;
  var cfg = null;
  var offset = 0;

  // ---------- Parameter, Speicher ----------
  function unb64(s) {
    try { return JSON.parse(decodeURIComponent(escape(atob(s)))); } catch (e) { return null; }
  }
  function b64(obj) { return btoa(unescape(encodeURIComponent(JSON.stringify(obj)))); }

  function store(kind, key, value) {
    try { (kind === 'local' ? localStorage : sessionStorage).setItem(key, JSON.stringify(value)); } catch (e) { /* gesperrt */ }
  }
  function load(kind, key) {
    try { var v = (kind === 'local' ? localStorage : sessionStorage).getItem(key); return v ? JSON.parse(v) : null; } catch (e) { return null; }
  }

  function readParams() {
    var params = {};
    location.hash.replace(/^#/, '').split('&').forEach(function (part) {
      var i = part.indexOf('=');
      if (i > 0) params[part.slice(0, i)] = decodeURIComponent(part.slice(i + 1));
    });
    if (params.c) {
      var c = unb64(params.c);
      if (c && c.u && c.k) { cfg = c; store('local', CFG_KEY, c); }
    }
    if (!cfg) cfg = load('local', CFG_KEY);
    offset = parseInt(params.o, 10) || 0;
    if (params.t) topic = unb64(params.t);
    if (topic) {
      var id = params.id || String(Date.now()) + Math.random().toString(16).slice(2, 6);
      store('session', TOPIC_PREFIX + id, topic);
      try { history.replaceState(null, '', location.pathname + '#id=' + id + '&o=' + offset + '&t=' + encodeURIComponent(b64(topic))); } catch (e) { /* egal */ }
    } else if (params.id) {
      topic = load('session', TOPIC_PREFIX + params.id);
    }
  }

  // ---------- Größe und Lage merken; offene Fenster zählen (für den Versatz) ----------
  var winId = String(Date.now()) + Math.random().toString(16).slice(2, 6);

  function heartbeat() {
    var open = load('local', OPEN_KEY) || {};
    var now = Date.now();
    Object.keys(open).forEach(function (k) { if (now - open[k] > 6000) delete open[k]; });
    open[winId] = now;
    store('local', OPEN_KEY, open);
  }
  function unregister() {
    var open = load('local', OPEN_KEY) || {};
    delete open[winId];
    store('local', OPEN_KEY, open);
  }

  var lastGeom = '';
  function remember() {
    // ohne eigenes Fenster (Tab) hat die Lage keine Bedeutung
    if (!window.outerWidth || window.outerWidth < 300) return;
    var g = { x: window.screenX - offset, y: window.screenY - offset, w: window.innerWidth, h: window.innerHeight };
    var s = JSON.stringify(g);
    if (s !== lastGeom) { lastGeom = s; store('local', GEOM_KEY, g); }
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
          if (r.status === 401) throw { code: 'UNAUTHORIZED', message: 'The XOB key is wrong. Check it in the settings of the panel.' };
          if (!r.ok || !json || json.status === 'error') {
            throw { code: (json && json.error && json.error.code) || ('HTTP_' + r.status), message: (json && json.error && json.error.message) || ('XOB answered with HTTP ' + r.status + '.') };
          }
          json._ms = Date.now() - started;
          return json;
        });
      }, function () {
        throw { code: 'NETWORK', message: 'XOB is not reachable. Check the network and try again.' };
      });
  }

  function search() {
    if (!topic) return Promise.resolve({ error: 'No search was handed over. Open this window from Outlook via xOrg.' });
    if (!cfg) return Promise.resolve({ error: 'The XOB access is missing. Open this window again from Outlook.' });
    return Core.relatedQuery(callXob, {
      conversationId: topic.conversationId || '',
      normalized: topic.normalized,
      useSubject: topic.useSubject,
      mode: topic.mode,
      terms: topic.terms || [],
      fullText: !!topic.fullText,
      top: topic.top || 50
    }).then(function (json) {
      var data = json.data || {};
      var result = Core.mergeResults(data.conversation || [], data.subject || [], {
        normalized: topic.normalized,
        prefixes: topic.prefixes || Core.DEFAULT_PREFIXES,
        currentInternetMessageId: topic.curImid,
        currentRestId: topic.curRest,
        currentMessage: topic.cur || null,
        maxResults: topic.maxResults || 25,
        useSubject: topic.useSubject,
        terms: topic.terms || [],
        fullText: !!topic.fullText,
        termHits: data.terms || []
      });
      return {
        ok: true, ms: json._ms, total: result.total,
        items: result.items.map(function (e) { var m = Object.assign({}, e.message); m.__current = !!e.current; m.__sources = e.sources; return m; })
      };
    }).catch(function (e) {
      return { error: (e && e.message) || 'Unknown error.' };
    });
  }

  window.related = {
    topic: function () {
      return Promise.resolve(topic ? { raw: topic.raw, normalized: topic.normalized, terms: topic.terms || [], fullText: !!topic.fullText, label: 'Related emails' } : null);
    },
    search: search,
    openMail: function (link) {
      if (!link) return Promise.resolve({ ok: false, error: 'No link available for this email.' });
      window.open(link, '_blank', 'noopener');
      return Promise.resolve({ ok: true });
    },
    onRefresh: function () { /* der Aktualisieren-Knopf genügt */ }
  };

  readParams();
  heartbeat();
  setInterval(function () { heartbeat(); remember(); }, 2000);
  window.addEventListener('resize', remember);
  window.addEventListener('beforeunload', function () { remember(); unregister(); });
})();
