/**
 * Der frueheste Zuhoerer der Seite. Klassisches Skript, kein Modul - mit Absicht.
 *
 * Ein `<script type="module">` wird ZURUECKGESTELLT: Es laeuft erst, wenn das Dokument
 * fertig gelesen ist. office.js dagegen laeuft sofort und stoesst von dort aus weitere
 * Ladevorgaenge an. Weist die Inhaltsrichtlinie einen davon ab, faellt die Meldung, bevor
 * ein Modul ueberhaupt existiert - und ist verloren. Genau das Ereignis wird hier
 * gebraucht, also muss der Zuhoerer VOR office.js haengen, und dafuer gibt es nur ein
 * klassisches Skript.
 *
 * Die Datei tut deshalb genau zwei Dinge und nichts sonst: Sie merkt sich den Zeitpunkt
 * null und schreibt abgewiesene Ladevorgaenge mit. `js/boot.js` liest beides spaeter aus.
 * Sie darf nicht wachsen - was hier scheitert, scheitert vor jeder Fehlerbehandlung.
 */
(function () {
  "use strict";

  /** Deckel. Ein Verstoss in einer Schleife fuellte sonst den Speicher. */
  var MAX = 12;

  var start = {
    /** @type {string[]} Abgewiesene Ladevorgaenge, ohne Wiederholungen. */
    csp: [],
    /** Zeitpunkt null, so frueh wie die Seite es zulaesst. */
    t0: typeof performance !== "undefined" && performance.now ? performance.now() : Date.now(),
  };

  window.__TANSS_START__ = start;

  document.addEventListener("securitypolicyviolation", function (event) {
    if (start.csp.length >= MAX) return;
    var directive = String(event.effectiveDirective || event.violatedDirective || "?");
    var blocked = String(event.blockedURI || "?").slice(0, 200);
    var entry = directive + " → " + blocked;
    if (start.csp.indexOf(entry) === -1) start.csp.push(entry);
  });
})();
