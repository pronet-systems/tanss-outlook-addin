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

import { T } from "../i18n/de.js";

/** Deckel. Ein Verstoss wiederholt sich in einer Schleife sonst bis zum Speicherende. */
const MAX_VIOLATIONS = 12;

/** @type {Array<[string, string]>} Geordnete Eintraege: Bezeichnung, Wert. */
const marks = [];

/**
 * Was `js/boot-early.js` bereits gesammelt hat.
 *
 * Jenes klassische Skript haengt VOR office.js und faengt damit auch die Verstoesse, die
 * waehrend dessen Start fallen. Fehlt es - weil jemand die Zeile aus der `index.html`
 * entfernt hat -, wird hier notduerftig selbst zugehoert; dann fehlt aber der Anfang, und
 * die Diagnoseseite sagt das, statt eine leere Liste als Entwarnung auszugeben.
 */
const early = globalThis.__TANSS_START__ || null;

/** @type {string[]} Blockierte Ladevorgaenge, in der Reihenfolge ihres Auftretens. */
const violations = early ? early.csp : [];

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
const t0 = early ? early.t0 : now();

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
if (!early && typeof document !== "undefined" && typeof document.addEventListener === "function") {
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
 * Draht zum Wirt gibt (`window.external`). Er traegt Host, Plattform, Version und Sprache
 * in einer mit `$` getrennten Kette. Fehlt er, weiss office.js nicht, welche
 * hostspezifische Datei es nachladen soll, und der Start kommt nie zum Ende. Genau das
 * laesst sich hier ablesen, statt es zu vermuten.
 */
export function hostInfoParam() {
  return extract(globalThis.location ? globalThis.location.search : "");
}

/**
 * Dieselbe Kennung, aber im ANKER gesucht - dort, wo office.js sie nie findet.
 *
 * Die Adressen im Manifest enden auf `#/ticket/neu`. Haengt der Wirt seine Kennung
 * schlicht hinten an die Zeichenkette, landet sie hinter dem Doppelkreuz und damit im
 * Anker. Dann steht sie zwar da, aber an der falschen Stelle, und der Start kommt nie zum
 * Ende. Ob es so kommt, entscheidet der Wirt; deshalb wird es nachgesehen und nicht
 * angenommen.
 */
export function hostInfoInHash() {
  return extract(globalThis.location ? globalThis.location.hash : "");
}

/** Sucht `_host_Info` in einem Adressteil, nach denselben Regeln wie office.js. */
function extract(part) {
  const text = String(part || "");
  if (!text) return "";
  const pieces = text.split("_host_Info=");
  if (pieces.length < 2) return "";
  const value = pieces[1].split(/[&#]/)[0];
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
}

/**
 * Die Zeile, die auf der Diagnoseseite UND unter einer Fehlermeldung steht.
 *
 * Sie wird an einer einzigen Stelle gebaut, damit Bildschirmfoto und kopierter Text nie
 * auseinanderlaufen - und damit der Fundort nicht an einem der beiden Orte verschwindet.
 */
export function hostInfoLine() {
  const inSearch = hostInfoParam();
  if (inSearch) return inSearch;
  const inHash = hostInfoInHash();
  if (inHash) return `${inHash} · ${T.diagnose.bootHostInfoInHash}`;
  return T.diagnose.bootHostInfoMissing;
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

/**
 * Ob der fruehe Zuhoerer lief.
 *
 * Ist er ausgefallen, ist eine leere Liste KEINE Entwarnung: Die Meldungen aus dem Start
 * von office.js waren dann schon durch, bevor ueberhaupt zugehoert wurde. Das gehoert
 * gesagt - sonst liest sich ein Loch wie ein Befund.
 */
export function hadEarlyListener() {
  return early !== null;
}

/** Alle festgehaltenen Messwerte, als Kopie. */
export function entries() {
  return marks.map((pair) => [pair[0], pair[1]]);
}
