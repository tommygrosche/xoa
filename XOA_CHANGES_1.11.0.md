# XOA 1.11.0 (2026-11-06)

**Fenster „Related emails“: Doppelklick öffnet, Klick wählt aus, neuer Knopf „Move Task Flag“** (gleich wie xOrg Desktop 2.20.0)

- Ein Klick (Tipp) wählt eine Mail aus oder ab, höchstens zwei; eine dritte ersetzt die zuerst gewählte. Doppelklick (Doppeltipp) öffnet die Mail im Browser. Eigene Doppelklick-Erkennung, funktioniert auf Windows und Android gleich.
- **Move Task Flag** (oben links neben der Suche): nur aktiv, wenn genau zwei Mails gewählt sind und genau eine davon ein Flag hat (gesetzt oder erledigt). Dann:
  1. die andere Mail wird geflaggt, Fälligkeit = Fälligkeit des alten Flags; liegt sie in der Vergangenheit, heute;
  2. danach wird das alte Flag komplett entfernt (nicht nur „erledigt“). Schlägt Schritt 1 fehl, bleibt das alte Flag unverändert.
  3. Die Aufgabe in To Do (Liste „Flagged email“) entsteht aus dem Flag selbst: To Do ersetzt sie beim nächsten Abgleich durch eine Aufgabe mit Link auf die neue Mail. Labels/Notizen der alten Aufgabe überträgt nur xOrg Desktop.
- Braucht Flow **xob-mail 1.4.0** (neuer Service `mail.flag`). Ohne ihn meldet der Knopf einen Fehler, die Suche läuft weiter.

Geänderte Dateien: `core.js` (gemeinsam mit xOrg `app/mailCore.js`), `viewer.js`, `viewer.css` (identisch mit xOrg `app/related.js`/`related.css`, enthalten jetzt auch den Desktop-Knopf „Link ToDo…“, der in XOA nie erscheint), `viewer.html`, `viewer-host.js`, `taskpane.js` (nur Versionsnummer).
Flach nach GitHub `tommygrosche/xoa` hochladen (ersetzen), ca. 10 Minuten Cache abwarten. Im Panel muss unten **v1.11.0** stehen; auf Android Outlook ganz schließen und neu öffnen.
