# XOA 1.11.1 (2026-11-06)

**Move Task Flag – Fälligkeit:** Hat das alte Flag keine Fälligkeit, bekommt das neue Flag jetzt ebenfalls **heute** (vorher ohne Datum). Regel wie in xOrg Desktop 2.20.1: keine oder in der Vergangenheit → heute, in der Zukunft → unverändert.

Geänderte Dateien: `core.js` (gemeinsam mit xOrg `app/mailCore.js`, Funktion `moveFlagDue`), `viewer-host.js` und `taskpane.js` (nur Versionsnummer).
Flach nach GitHub `tommygrosche/xoa` hochladen (ersetzen). Im Panel muss unten **v1.11.1** stehen. Kein Flow-Update nötig.
