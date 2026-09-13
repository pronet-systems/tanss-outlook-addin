/**
 * Das Startprotokoll: was beim Hochfahren des Panes tatsaechlich passiert ist.
 *
 * Warum es das gibt: Im klassischen Outlook laeuft das Taskpane in einem eingebetteten
 * Browser OHNE Entwicklerwerkzeuge. Bleibt der Start haengen, gibt es keine Konsole, in die
 * jemand sehen koennte - und damit keinen Unterschied zwischen "office.js wurde gar nicht
 * geladen", "die Inhaltsrichtlinie hat etwas blockiert" und "Outlook hat sich einfach nicht
 * gemeldet". Das sind drei voellig verschiedene Ursachen mit drei verschiedenen Loesungen,
 * und ohne Messung ist die Wahl zwischen ihnen geraten.
 *
 * Dieses Modul misst deshalb genau die vier Dinge, die den Fall entscheiden, und die
 * Diagnoseseite zeigt sie an. Ein Bildschirmfoto des Panes beantwortet die Frage dann,
 * ohne dass jemand einen zweiten Client aufsetzen muss.
 *
 * Es MUSS als erstes Modul geladen werden - `main.js` importiert es vor allem anderen.
 * Verstoesse gegen die Inhaltsrichtlinie werden nur ab dem Zeitpunkt gemeldet, an dem der
 * Zuhoerer haengt; alles davor ist verloren. Das betrifft genau einen Fall, naemlich den
 * Skriptknoten von office.js selbst - und der ist auf anderem Weg messbar, siehe
 * `officePresent`.
 *
 * Was hier NICHT hineingeschrieben wird: nichts aus einer Mail, kein Token, kein Kennwort.
 * Das Protokoll wandert ueber "Als Text kopieren" in fremde Haende.
 */

/** Deckel. Ein Verstoss wiederholt sich in einer Schleife sonst bis zum Speicherende. */
const MAX_VIOLATIONS = 12;

/** @type {Array<[string, string]>} Geordnete Eintraege: Bezeichnung, Wert. */
const marks = [];

/** @type {string[]} Blockierte Ladevorgaenge, in der Reihenfolge ihres Auftretens. */
const violations = [];

/**
 * War die Office-Bibliothek da, als das erste eigene Modul lief?
 *
 * Der Wert wird SOFORT festgehalten und nicht spaeter erfragt. office.js laedt synchron im
 * Kopf der Seite, also vor jedem Modul: Fehlt `window.Office` hier, ist der Skriptknoten
 * gescheitert - falsche Adresse, kein Netz, oder von der Inhaltsrichtlinie abgewiesen. Eine
 * spaetere Abfrage koennte das nicht mehr von "geladen, aber stumm" unterscheiden.
 */
const officePresent = typeof globalThis.Office === "object" && globalThis.Office !== null;

/**
 * Der Startzeitpunkt, gegen den alle Dauern gemessen werden.
 *
 * `performance.now()`, nicht `Date.now()`: Eine Zeitumstellung oder ein Abgleich der
 * Systemuhr mitten im Start ergaebe sonst negative Dauern.
 */
const t0 = now();

function now() {
  const p = globalThis.performance;
  return p && typeof p.now === "function" ? p.now() : Date.now();
}

/** Millisekunden seit dem Start des Panes, ganzzahlig. */
export function sinceStart() {
  return Math.round(now() - t0);
}

/**
 * Verstoesse gegen die Inhaltsrichtlinie mitschreiben.
 *
 * Genau zwei Angaben je Verstoss: welche Direktive gegriffen hat und was sie abgewiesen hat.
 * Beides zusammen benennt die Zeile in `index.html`, die fehlt - mehr braucht niemand, und
 * der Rest des Ereignisses (Zeilennummer, Beispieltext) wuerde nur die Liste fluten.
 */
if (typeof document !== "undefined" && typeof document.addEventListener === "function") {
  document.addEventListener("securitypolicyviolation", (event) => {
    if (violations.length >= MAX_VIOLATIONS) return;
    const directive = String(event.effectiveDirective || event.violatedDirective || "?");
    const blocked = String(event.blockedURI || "?").slice(0, 200);
    const entry = `${directive} → ${blocked}`;
    // Dieselbe Datei wird beim Laden oft mehrfach angefragt. Einmal nennen genuegt.
    if (!violations.includes(entry)) violations.push(entry);
  });
}

/**
 * Der Wert `_host_Info`, den Outlook an die Seitenadresse haengt.
 *
 * office.js liest ihn aus `window.location.search` - und NUR von dort, wenn es keinen
 * Draht zum Wirt gibt (`window.external`). Er traegt Host, Plattform, Fassung und Sprache
 * in einer mit `$` getrennten Kette. Fehlt er, weiss office.js nicht, welche
 * hostspezifische Datei es nachladen soll, und der Start kommt nie zum Ende. Genau das
 * laesst sich hier ablesen, statt es zu vermuten.
 */
export function hostInfoParam() {
  const search = globalThis.location ? String(globalThis.location.search || "") : "";
  if (!search) return "";
  const parts = search.split("_host_Info=");
  if (parts.length < 2) return "";
  const value = parts[1].split(/[&#]/)[0];
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
}

/** Einen Messwert festhalten. Spaetere Aufrufe mit demselben Namen ueberschreiben. */
export function mark(label, value) {
  const index = marks.findIndex(([name]) => name === label);
  const entry = [label, String(value)];
  if (index >= 0) marks[index] = entry;
  else marks.push(entry);
}

/** Ob die Office-Bibliothek beim Start der eigenen Module vorhanden war. */
export function hadOfficeLibrary() {
  return officePresent;
}

/** Die blockierten Ladevorgaenge, als Kopie. */
export function blocked() {
  return violations.slice();
}

/** Alle festgehaltenen Messwerte, als Kopie. */
export function entries() {
  return marks.map((pair) => [pair[0], pair[1]]);
}
