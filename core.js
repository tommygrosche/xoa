/*
 * xoa core - reine Logik ohne Office.js (testbar)
 * Betreff-Normalisierung, Suchbegriff, Zusammenführen der xob-Treffer, Formatierung
 */
(function (root) {
  'use strict';

  var DEFAULT_PREFIXES = ['RE:', 'AW:', 'FW:', 'FWD:', 'WG:', 'Antw:', 'Antwort:', 'Weitergeleitet:', 'TR:'];

  function escapeRe(s) {
    return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  function cleanPrefixList(prefixes) {
    var seen = {};
    var out = [];
    (prefixes || []).forEach(function (p) {
      var t = String(p || '').replace(/\s+/g, ' ').trim();
      if (!t) return;
      var k = t.toLowerCase();
      if (seen[k]) return;
      seen[k] = true;
      out.push(t);
    });
    return out;
  }

  function buildPrefixRegexes(prefixes) {
    return cleanPrefixList(prefixes)
      .sort(function (a, b) { return b.length - a.length; })
      .map(function (p) {
        // Endet der Präfix auf Buchstabe/Ziffer (z. B. "WG"), muss danach Leerzeichen, ":" oder Ende folgen,
        // damit "WGT Meeting" nicht zu "T Meeting" wird.
        var needsBoundary = /[A-Za-z0-9ÄÖÜäöüß]$/.test(p);
        return new RegExp('^' + escapeRe(p) + (needsBoundary ? '(?=[\\s:]|$)' : '') + '[\\s:]*', 'i');
      });
  }

  // Nummerierte Antworten wie "RE[2]:" oder "AW(3):"
  var NUMBERED = /^[A-Za-zÄÖÜäöü]{1,4}\s*[\[(]\d+[\])]\s*:\s*/;

  function normalizeSubject(subject, prefixes, regexCache) {
    var s = String(subject || '').replace(/\s+/g, ' ').trim();
    var res = regexCache || buildPrefixRegexes(prefixes);
    var changed = true;
    var guard = 0;
    while (changed && s && guard < 50) {
      changed = false;
      guard++;
      for (var i = 0; i < res.length; i++) {
        if (res[i].test(s)) {
          s = s.replace(res[i], '').trim();
          changed = true;
        }
      }
      if (NUMBERED.test(s)) {
        s = s.replace(NUMBERED, '').trim();
        changed = true;
      }
    }
    return s;
  }

  function compareKey(normalized) {
    return String(normalized || '').replace(/\s+/g, ' ').trim().toLowerCase();
  }

  // KQL-Suchwert für Graph $search. mode: "subject" (nur Betreff) oder "all" (Betreff, Text, Absender)
  function buildSearch(normalized, mode) {
    var t = String(normalized || '').replace(/["\\]/g, ' ').replace(/\s+/g, ' ').trim();
    if (!t) return '';
    if (mode === 'all') return '"' + t + '"';
    return '"subject:\\"' + t + '\\""';
  }

  function messageDate(m) {
    var d = m && (m.receivedDateTime || m.sentDateTime);
    var t = d ? Date.parse(d) : NaN;
    return isNaN(t) ? 0 : t;
  }

  /*
   * Führt Conversation- und Betreff-Treffer zusammen. Die ausgewählte Mail ist immer enthalten
   * (entry.current = true); fehlt sie in den Treffern, wird opts.currentMessage eingesetzt.
   * opts: { normalized, prefixes, currentInternetMessageId, currentRestId, currentMessage, maxResults, useSubject }
   * Rückgabe: { items, total }
   */
  function mergeResults(conversation, subject, opts) {
    var res = buildPrefixRegexes(opts.prefixes);
    var target = compareKey(opts.normalized);
    var byKey = {};
    var list = [];
    var current = null;

    function isCurrent(m) {
      return (opts.currentInternetMessageId && m.internetMessageId === opts.currentInternetMessageId) ||
        (opts.currentRestId && m.id === opts.currentRestId);
    }

    function add(m, source) {
      if (!m || m.isDraft) return;
      var key = m.internetMessageId || m.id;
      if (!key) return;
      var existing = byKey[key];
      if (existing) {
        if (existing.sources.indexOf(source) < 0) existing.sources.push(source);
        return;
      }
      var entry = { message: m, sources: [source], current: !!isCurrent(m) };
      if (entry.current) {
        if (current) return; // gleiche Mail in zweitem Ordner
        current = entry;
      }
      byKey[key] = entry;
      list.push(entry);
    }

    (conversation || []).forEach(function (m) { add(m, 'conversation'); });
    if (opts.useSubject && target) {
      (subject || []).forEach(function (m) {
        if (compareKey(normalizeSubject(m && m.subject, null, res)) === target) add(m, 'subject');
      });
    }
    if (!current && opts.currentMessage) {
      current = { message: opts.currentMessage, sources: ['current'], current: true };
      list.push(current);
    }

    list.sort(function (a, b) { return messageDate(b.message) - messageDate(a.message); });
    var max = Math.max(1, parseInt(opts.maxResults, 10) || 25);
    var items = list.slice(0, max);
    if (current && items.indexOf(current) < 0) {
      items[items.length - 1] = current;
      items.sort(function (a, b) { return messageDate(b.message) - messageDate(a.message); });
    }
    return { items: items, total: list.length };
  }

  function pad(n) { return (n < 10 ? '0' : '') + n; }

  function formatTimestamp(iso) {
    var t = iso ? new Date(iso) : null;
    if (!t || isNaN(t.getTime())) return '';
    return pad(t.getDate()) + '.' + pad(t.getMonth() + 1) + '.' + t.getFullYear() + ' ' + pad(t.getHours()) + ':' + pad(t.getMinutes());
  }

  function personName(r) {
    var e = r && r.emailAddress;
    if (!e) return '';
    return e.name || e.address || '';
  }

  function recipientsLabel(list) {
    var arr = (list || []).filter(Boolean);
    if (!arr.length) return '';
    var first = personName(arr[0]);
    return arr.length > 1 ? first + ' +' + (arr.length - 1) : first;
  }

  function flagState(m) {
    var s = m && m.flag && m.flag.flagStatus;
    if (s === 'flagged') return 'flagged';
    if (s === 'complete') return 'complete';
    return 'none';
  }

  var api = {
    DEFAULT_PREFIXES: DEFAULT_PREFIXES,
    cleanPrefixList: cleanPrefixList,
    buildPrefixRegexes: buildPrefixRegexes,
    normalizeSubject: normalizeSubject,
    compareKey: compareKey,
    buildSearch: buildSearch,
    mergeResults: mergeResults,
    formatTimestamp: formatTimestamp,
    personName: personName,
    recipientsLabel: recipientsLabel,
    flagState: flagState
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.XoaCore = api;
})(this);
