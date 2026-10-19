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

  // "Projekt 1; Kunde 1" -> ["Projekt 1", "Kunde 1"] (ohne Leere/Doppelte, max. 10)
  function parseTerms(text) {
    var seen = {};
    var out = [];
    String(text || '').split(';').forEach(function (t) {
      var v = t.replace(/["\\]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 100);
      if (!v) return;
      var k = v.toLowerCase();
      if (seen[k]) return;
      seen[k] = true;
      out.push(v);
    });
    return out.slice(0, 10);
  }

  // KQL für Graph $search: jeder Begriff als ganze Phrase im Betreff, verknüpft mit OR
  function buildTermsSearch(terms) {
    var list = (terms || []).filter(Boolean);
    if (!list.length) return '';
    return '"' + list.map(function (t) { return 'subject:\\"' + t + '\\"'; }).join(' OR ') + '"';
  }

  // Enthält der Betreff mindestens einen Begriff vollständig (als ganze Phrase, nicht als Wortteil)?
  // "Projekt 1" passt auf "RE: Projekt 1 Kickoff", aber nicht auf "Projekt 12".
  function subjectHasTerm(subject, terms) {
    var s = String(subject || '').replace(/\s+/g, ' ');
    return (terms || []).some(function (t) {
      var v = String(t || '').replace(/\s+/g, ' ').trim();
      if (!v) return false;
      var re = new RegExp('(^|[^\\p{L}\\p{N}])' + escapeRe(v) + '(?=$|[^\\p{L}\\p{N}])', 'iu');
      return re.test(s);
    });
  }

  /*
   * Fragt xob mail.related ab. call(service, payload) -> Promise({data, _ms}).
   * p: { conversationId, normalized, useSubject, mode, terms[], top }
   * Fällt bei abgelehnter Abfrage auf allgemeine Betreffsuche und einzelne Begriffe zurück.
   */
  function relatedQuery(call, p) {
    var terms = p.terms || [];
    function payload(mode, withTerms) {
      return {
        conversationId: p.conversationId || '',
        search: p.useSubject ? buildSearch(p.normalized, mode) : '',
        terms: withTerms ? buildTermsSearch(terms) : '',
        top: p.top || 50
      };
    }
    var mode = p.mode === 'all' ? 'all' : 'subject';
    return call('mail.related', payload(mode, true)).catch(function (e) {
      if (e && (e.code === 'UNAUTHORIZED' || e.code === 'NETWORK' || e.code === 'TIMEOUT')) throw e;
      var calls = [call('mail.related', payload('all', false))];
      terms.forEach(function (t) {
        calls.push(call('mail.related', { conversationId: '', search: '', terms: buildTermsSearch([t]), top: p.top || 50 }));
      });
      return Promise.all(calls).then(function (rs) {
        var d = { conversation: [], subject: [], terms: [] };
        var ms = 0;
        rs.forEach(function (r) {
          var x = (r && r.data) || {};
          d.conversation = d.conversation.concat(x.conversation || []);
          d.subject = d.subject.concat(x.subject || []);
          d.terms = d.terms.concat(x.terms || []);
          ms = Math.max(ms, (r && r._ms) || 0);
        });
        return { data: d, _ms: ms };
      });
    });
  }

  function messageDate(m) {
    var d = m && (m.receivedDateTime || m.sentDateTime);
    var t = d ? Date.parse(d) : NaN;
    return isNaN(t) ? 0 : t;
  }

  /*
   * Führt Conversation- und Betreff-Treffer zusammen. Die ausgewählte Mail ist immer enthalten
   * (entry.current = true); fehlt sie in den Treffern, wird opts.currentMessage eingesetzt.
   * opts: { normalized, prefixes, currentInternetMessageId, currentRestId, currentMessage, maxResults, useSubject, terms, termHits }
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
    if (opts.terms && opts.terms.length) {
      (opts.termHits || []).forEach(function (m) {
        if (m && subjectHasTerm(m.subject, opts.terms)) add(m, 'term');
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
    parseTerms: parseTerms,
    buildTermsSearch: buildTermsSearch,
    subjectHasTerm: subjectHasTerm,
    relatedQuery: relatedQuery,
    mergeResults: mergeResults,
    formatTimestamp: formatTimestamp,
    personName: personName,
    recipientsLabel: recipientsLabel,
    flagState: flagState
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.XoaCore = api;
})(this);
