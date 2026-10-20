# XOA 1.5.0

- Fenster „Related emails“ (`viewer.html`, `viewer.css`, `viewer.js`) ist jetzt **derselbe Code wie in xOrg Desktop** (nur `viewer-host.js` ist XOA-spezifisch: Suche gegen xob, Mail öffnen, Größe/Lage merken). Änderungen am Fenster gelten künftig für beide Apps.
- Look wie xOrg, Spalten verschiebbar, Sortieren per Klick auf den Titel, Freitext-Filter je Spalte und über alles, zweite Zeile je Mail (250 Zeichen), feine Trennlinie, Schrift und Icons skalieren mit der Fensterbreite, Fenstertitel = Betreff der Ausgangsmail.
- Das Panel öffnet das Fenster jetzt mit `window.open` als eigenes Popup (Größe und Position aus dem letzten Fenster, weitere Fenster 30 px nach rechts unten versetzt). Wird das Popup blockiert, öffnet wie bisher `openBrowserWindow` (dann ohne gemerkte Größe/Lage).
- Mail öffnen: ein Klick (statt Doppelklick).
- From/To zeigen nur die Adresse.
- Der Mailtext (zweite Zeile) kommt aus dem Feld `bodyPreview`; dafür den Flow `xob-mail` auf 1.3.0 aktualisieren (`xob-mail-1.3.0-update.zip`).
- Dateien: `viewer.html`, `viewer.css`, `viewer.js`, `viewer-host.js` (neu), `taskpane.js` (VERSION 1.5.0), `core.js` unverändert. `logo.png` wird wie bisher aus dem Ordner geladen. In `manifest.xml` ggf. die Version auf 1.5.0 setzen und `viewer-host.js` mit ausliefern (Cache beachten).

# XOA 1.5.1

- `core.js` und `viewer.js` an xOrg Desktop 2.10.0 angepasst (Unterstützung für Volltextsuche `fullText`; das Fenster nennt, ob Betreff- oder Volltextsuche lief). Im Panel gibt es den Schalter „Full text“ noch nicht.
- Ersetze `core.js`, `viewer.js` und `taskpane.js` (nur die Versionsnummer).

# XOA 1.6.0

- **Schalter „Volltext“** unter dem Suchfeld (wird vom Skript eingefügt, `taskpane.html` bleibt unverändert). Aus: Suchtext nur im Betreff. An: auch im Mailtext. Die Wahl wird im Browser gemerkt; ohne Wahl gilt die Einstellung „Suchmodus“ (`all` = Volltext). Umschalten startet die Suche neu.
- Das Fenster „Related emails“ bekommt die Wahl mit und nennt oben, ob eine Betreff- oder Volltextsuche lief.
- Geänderte Dateien: `taskpane.js`, `viewer-host.js`, `core.js`, `viewer.js` (alle im Paket).
