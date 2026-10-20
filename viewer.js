/*
 * xOrg / XOA - Fenster "Related emails" (gemeinsame Oberfläche)
 * Dieselbe Datei läuft in xOrg Desktop (related.js) und in XOA (viewer.js).
 * Die App stellt window.related bereit:
 *   topic()            -> Promise<{ raw, normalized, terms, label }>
 *   search()           -> Promise<{ items, total, ms } | { error }>
 *   openMail(link, id) -> Promise<{ ok, error? }>
 *   openOn             optional: 'dblclick' = Mail erst nach Doppelklick öffnen (Standard: ein Klick)
 *   onRefresh(cb)      Aufruf, wenn das Fenster neu suchen soll
 */
(function () {
  'use strict';

  var Core = window.XoaCore;
  var $ = function (id) { return document.getElementById(id); };

  var COLS = [
    { key: 'from', title: 'From', w: 19, text: true },
    { key: 'to', title: 'To', w: 21, text: true },
    { key: 'subject', title: 'Subject', w: 34, text: true },
    { key: 'received', title: 'Received', w: 14, text: true },
    { key: 'hasAttachments', title: 'Attachment', w: 6, icon: 'att' },
    { key: 'flag', title: 'Flag', w: 6, icon: 'flag' }
  ];
  var MIN_W = 4;
  var WIDTH_KEY = 'xorgRelatedCols';
  var SNIPPET = 250;

  var topic = null;
  var view = { kind: 'loading' };
  var sortKey = 'received';
  var sortDir = -1;
  var selectedId = null;
  var colFilter = { from: '', to: '', subject: '', received: '', hasAttachments: false, flag: false };
  var widths = COLS.map(function (c) { return c.w; });

  var CLIP = '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M10.5 4.5 5.8 9.2a1.4 1.4 0 0 0 2 2l5-5a2.8 2.8 0 0 0-4-4l-5.2 5.2a4.2 4.2 0 0 0 6 6l4.2-4.2" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"/></svg>';
  var FLAG = '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M3.5 14.5V2.5m0 0h8l-1.8 3 1.8 3h-8" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>';
  var DONE = '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M3 8.5l3 3 7-7" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>';

  // ---------- Speicher (nur Spaltenbreiten; Fenstergröße/-lage macht die App) ----------
  function loadWidths() {
    try {
      var v = JSON.parse(localStorage.getItem(WIDTH_KEY) || 'null');
      if (v && v.length === COLS.length && v.every(function (n) { return typeof n === 'number' && n >= MIN_W; })) widths = v;
    } catch (e) { /* Speicher gesperrt */ }
  }
  function saveWidths() {
    try { localStorage.setItem(WIDTH_KEY, JSON.stringify(widths)); } catch (e) { /* egal */ }
  }

  // ---------- Skalierung: die Breite bestimmt die Größe, die Höhe die Anzahl der Mails ----------
  function scale() {
    var px = Math.max(10, Math.min(24, window.innerWidth / 85));
    document.documentElement.style.fontSize = px.toFixed(2) + 'px';
  }

  // ---------- Suche ----------
  function search() {
    view = { kind: 'loading' };
    render();
    window.related.search().then(function (r) {
      if (!r || r.error) {
        view = { kind: 'error', text: (r && r.error) || 'Unknown error.' };
        render();
        return;
      }
      view = {
        kind: 'results', ms: r.ms,
        count: r.total > r.items.length ? r.items.length + ' of ' + r.total : String(r.items.length),
        total: r.total,
        items: r.items
      };
      render();
    }).catch(function (e) {
      view = { kind: 'error', text: (e && e.message) || 'Unknown error.' };
      render();
    });
  }

  function setState(text, isError) {
    var s = $('state');
    s.textContent = text || '';
    s.className = 'state' + (isError ? ' error' : '');
  }

  function addrOf(r) { var e = r && r.emailAddress; return e ? (e.address || e.name || '') : ''; }

  function snippet(m) {
    var t = Core.stripBanners(String(m.bodyPreview || m.__text || ''), topic && topic.banners).replace(/\s+/g, ' ').trim();
    return t.length > SNIPPET ? t.slice(0, SNIPPET).replace(/\s+\S*$/, '') + '…' : t;
  }

  function fields(m) {
    var received = m.receivedDateTime || m.sentDateTime || '';
    return {
      id: m.id,
      webLink: m.webLink || '',
      from: addrOf(m.from),
      to: (m.toRecipients || []).map(addrOf).filter(Boolean).join('; '),
      subject: m.subject || '',
      received: received,
      receivedText: Core.formatTimestamp(received),
      text: snippet(m),
      hasAttachments: !!m.hasAttachments,
      flag: Core.flagState(m),
      current: !!m.__current,
      sources: m.__sources || []
    };
  }

  function sortValue(it, key) {
    if (key === 'received') { var t = Date.parse(it.received); return isNaN(t) ? 0 : t; }
    if (key === 'hasAttachments') return it.hasAttachments ? 1 : 0;
    if (key === 'flag') return it.flag === 'flagged' ? 2 : (it.flag === 'complete' ? 1 : 0);
    return String(it[key] || '').toLowerCase();
  }

  function has(hay, needle) { return String(hay || '').toLowerCase().indexOf(needle) >= 0; }

  function visibleItems() {
    var items = (view.items || []).map(fields);
    var q = $('filter').value.trim().toLowerCase();
    items = items.filter(function (it) {
      if (q && !has([it.from, it.to, it.subject, it.receivedText, it.text].join(' '), q)) return false;
      if (colFilter.from && !has(it.from, colFilter.from)) return false;
      if (colFilter.to && !has(it.to, colFilter.to)) return false;
      if (colFilter.subject && !has(it.subject, colFilter.subject)) return false;
      if (colFilter.received && !has(it.receivedText, colFilter.received)) return false;
      if (colFilter.hasAttachments && !it.hasAttachments) return false;
      if (colFilter.flag && it.flag === 'none') return false;
      return true;
    });
    items.sort(function (a, b) {
      var x = sortValue(a, sortKey), y = sortValue(b, sortKey);
      if (x < y) return -1 * sortDir;
      if (x > y) return 1 * sortDir;
      return sortValue(b, 'received') - sortValue(a, 'received');
    });
    return items;
  }

  function openItem(it) {
    window.related.openMail(it.webLink, it.id).then(function (r) { if (r && !r.ok) setState(r.error || 'Outlook could not be opened.', true); });
  }

  function cell(tag, cls, text, title) {
    var c = document.createElement(tag);
    c.className = cls;
    c.textContent = text;
    if (title) c.title = title;
    return c;
  }

  function select(group, it) {
    var prev = document.querySelector('tbody.sel');
    if (prev && prev !== group) prev.classList.remove('sel');
    group.classList.add('sel');
    selectedId = it.id;
  }

  // ---------- Darstellung ----------
  function render() {
    var raw = (topic && topic.raw) || '';
    var subj = raw || (topic && topic.normalized) || '–';
    $('subject').textContent = subj;
    $('subject').title = subj;
    var terms = (topic && topic.terms) || [];
    $('terms').textContent = terms.length ? (topic.fullText ? 'Full text search: ' : 'Subject search: ') + terms.join('; ') : '';
    document.title = subj;

    var wrap = $('grid');
    Array.prototype.slice.call(wrap.querySelectorAll('tbody')).forEach(function (b) { wrap.removeChild(b); });
    $('count').textContent = '';
    $('meta').textContent = '';

    if (view.kind === 'loading') { setState('Searching…'); return; }
    if (view.kind === 'message') { setState(view.text || ''); return; }
    if (view.kind === 'error') { setState(view.text || 'Unknown error.', true); return; }

    var items = visibleItems();
    var total = (view.items || []).length;
    $('count').textContent = items.length === total ? (view.count || String(total)) : items.length + ' of ' + total;
    $('meta').textContent = 'XOB · ' + view.ms + ' ms' + (window.related.info ? ' · ' + window.related.info() : '');

    if (!total) { setState('No emails found.'); return; }
    if (!items.length) { setState('No match for this filter.'); return; }
    setState('');

    var SRC = { conversation: 'same conversation', subject: 'same subject', term: 'search text', current: 'task email' };

    items.forEach(function (it) {
      var group = document.createElement('tbody');
      group.className = 'mail' + (it.current ? ' current' : '');
      group.tabIndex = 0;
      group.title = 'Found by: ' + it.sources.map(function (k) { return SRC[k] || k; }).join(', ');

      var head = document.createElement('tr');
      head.className = 'head';
      var from = cell('td', 'from', '', it.from);
      if (it.current) {
        var badge = document.createElement('span');
        badge.className = 'cur-badge';
        badge.textContent = 'Task email';
        from.appendChild(badge);
      }
      from.appendChild(document.createTextNode(it.from || '(unknown)'));
      head.appendChild(from);
      head.appendChild(cell('td', 'to', it.to, it.to));
      head.appendChild(cell('td', 'subject', it.subject || '(no subject)', it.subject));
      head.appendChild(cell('td', 'received', it.receivedText));
      var a = cell('td', 'ic att', '');
      if (it.hasAttachments) { a.innerHTML = CLIP; a.title = 'With attachment'; }
      head.appendChild(a);
      var f = cell('td', 'ic flag ' + it.flag, '');
      if (it.flag === 'flagged') { f.innerHTML = FLAG; f.title = 'Flagged for follow-up'; }
      else if (it.flag === 'complete') { f.innerHTML = DONE; f.title = 'Completed'; }
      head.appendChild(f);
      group.appendChild(head);

      var textRow = document.createElement('tr');
      textRow.className = 'text';
      var tc = document.createElement('td');
      tc.colSpan = COLS.length;
      var sn = document.createElement('div');
      sn.className = 'snip';
      sn.textContent = it.text;
      tc.appendChild(sn);
      textRow.appendChild(tc);
      group.appendChild(textRow);

      // Standard (xOrg Desktop): ein Klick öffnet. Setzt die App window.related.openOn = 'dblclick' (XOA), öffnet erst der Doppelklick.
      var dbl = window.related.openOn === 'dblclick';
      group.addEventListener('click', function () { select(group, it); if (!dbl) openItem(it); });
      if (dbl) group.addEventListener('dblclick', function () { select(group, it); openItem(it); });
      group.addEventListener('keydown', function (ev) {
        if (ev.key === 'Enter') { ev.preventDefault(); openItem(it); }
        else if (ev.key === 'ArrowDown' && group.nextSibling) { ev.preventDefault(); group.nextSibling.focus(); }
        else if (ev.key === 'ArrowUp' && group.previousSibling && group.previousSibling.tagName === 'TBODY') { ev.preventDefault(); group.previousSibling.focus(); }
      });
      group.addEventListener('focus', function () { select(group, it); });
      if (it.id === selectedId) group.classList.add('sel');
      $('grid').appendChild(group);
    });
  }

  // ---------- Kopf: Sortieren, Breite ziehen, Filter ----------
  function applyWidths() {
    var cols = $('cols').children;
    for (var i = 0; i < cols.length; i++) cols[i].style.width = widths[i] + '%';
  }

  function buildHead() {
    var cols = $('cols');
    var heads = $('heads');
    var filters = $('filters');
    COLS.forEach(function (c, i) {
      cols.appendChild(document.createElement('col'));

      var th = document.createElement('th');
      th.setAttribute('data-key', c.key);
      th.scope = 'col';
      if (c.icon) th.className = 'ic';
      if (c.key === sortKey) th.setAttribute('aria-sort', sortDir === 1 ? 'ascending' : 'descending');
      var box = document.createElement('div');
      box.className = 'hd';
      var b = document.createElement('button');
      b.type = 'button';
      b.className = 'sort';
      b.title = 'Click to sort by ' + c.title.toLowerCase();
      if (c.icon) { b.innerHTML = c.icon === 'att' ? CLIP : FLAG; b.setAttribute('aria-label', c.title); }
      else b.textContent = c.title;
      b.addEventListener('click', function () {
        if (sortKey === c.key) sortDir = -sortDir;
        else { sortKey = c.key; sortDir = (c.key === 'received' || c.icon) ? -1 : 1; }
        Array.prototype.forEach.call(heads.children, function (o) { o.removeAttribute('aria-sort'); });
        th.setAttribute('aria-sort', sortDir === 1 ? 'ascending' : 'descending');
        render();
      });
      box.appendChild(b);
      if (i < COLS.length - 1) box.appendChild(grip(i));
      th.appendChild(box);
      heads.appendChild(th);

      var fth = document.createElement('th');
      if (c.icon) {
        var t = document.createElement('button');
        t.type = 'button';
        t.className = 'tog';
        t.innerHTML = c.icon === 'att' ? CLIP : FLAG;
        t.title = c.icon === 'att' ? 'Show only emails with attachment' : 'Show only flagged or completed emails';
        t.addEventListener('click', function () {
          colFilter[c.key] = !colFilter[c.key];
          t.classList.toggle('on', colFilter[c.key]);
          render();
        });
        fth.appendChild(t);
      } else {
        var inp = document.createElement('input');
        inp.type = 'text';
        inp.spellcheck = false;
        inp.placeholder = 'Filter';
        inp.title = 'Free text filter for ' + c.title.toLowerCase();
        inp.addEventListener('input', function () { colFilter[c.key] = inp.value.trim().toLowerCase(); render(); });
        fth.appendChild(inp);
      }
      filters.appendChild(fth);
    });
    applyWidths();
  }

  function grip(i) {
    var g = document.createElement('span');
    g.className = 'grip';
    g.title = 'Drag to resize the column';
    g.addEventListener('click', function (e) { e.stopPropagation(); });
    g.addEventListener('mousedown', function (e) {
      e.preventDefault();
      e.stopPropagation();
      var total = $('grid').getBoundingClientRect().width;
      var startX = e.clientX;
      var a0 = widths[i], b0 = widths[i + 1];
      g.classList.add('drag');
      function move(ev) {
        var d = (ev.clientX - startX) / total * 100;
        d = Math.max(MIN_W - a0, Math.min(b0 - MIN_W, d));
        widths[i] = a0 + d;
        widths[i + 1] = b0 - d;
        applyWidths();
      }
      function up() {
        g.classList.remove('drag');
        document.removeEventListener('mousemove', move);
        document.removeEventListener('mouseup', up);
        saveWidths();
      }
      document.addEventListener('mousemove', move);
      document.addEventListener('mouseup', up);
    });
    return g;
  }

  function start() {
    loadWidths();
    buildHead();
    scale();
    window.addEventListener('resize', scale);
    $('filter').addEventListener('input', render);
    $('btnRefresh').addEventListener('click', search);
    window.related.onRefresh(search);
    $('opened').textContent = 'opened ' + Core.formatTimestamp(new Date().toISOString());
    window.related.topic().then(function (t) {
      topic = t;
      if (!topic) { view = { kind: 'message', text: 'No search was handed over. Open this window from a task in xOrg.' }; render(); return; }
      if (topic.label) $('kind').textContent = topic.label;
      search();
    });
  }

  start();
})();
