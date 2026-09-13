/**
 * Das Startprotokoll ist das Werkzeug fuer den Fall, in dem es kein anderes gibt: Im
 * klassischen Outlook laeuft das Taskpane ohne Entwicklerwerkzeuge, und bleibt der Start
 * haengen, ist die Diagnoseseite die einzige Quelle. Ein Protokoll, das in genau diesem
 * Moment selbst falsch liegt, waere schlimmer als keines - es schickte die Suche in die
 * falsche Richtung.
 *
 * Die Faelle hier decken die Auswertung von `_host_Info` ab. Der Wert entscheidet, ob
 * office.js ueberhaupt weiss, welchen Wirt es vor sich hat; er kommt als Parameter in der
 * Seitenadresse, und das Pane liest ihn nach denselben Regeln wie office.js selbst.
 */

import test from "node:test";
import assert from "node:assert/strict";

import { entries, hostInfoParam, mark, sinceStart } from "../../taskpane/js/boot.js";

/** Setzt die Adresse, die das Modul liest. Es fragt sie bei jedem Aufruf neu ab. */
function adresse(search) {
  globalThis.location = { search };
}

test("ohne Parameter meldet das Protokoll einen leeren Wert, nicht undefined", () => {
  adresse("?tanss=https%3A%2F%2Ftanss.example.org");
  assert.equal(hostInfoParam(), "");
});

test("ganz ohne Fragezeichen meldet es ebenfalls einen leeren Wert", () => {
  adresse("");
  assert.equal(hostInfoParam(), "");
});

test("die Kennung wird gefunden, auch wenn eigene Parameter davor stehen", () => {
  // Genau der Fall dieser Auslieferung: Das Manifest bringt `tanss` und `entra` mit,
  // Outlook haengt seine Kennung hinten an.
  adresse("?tanss=https%3A%2F%2Ftanss.example.org&entra=abc&_host_Info=Outlook%24Win32%2416.02%24de-de%24");
  assert.equal(hostInfoParam(), "Outlook$Win32$16.02$de-de$");
});

test("der Wert endet am naechsten Parameter und nicht erst am Zeilenende", () => {
  adresse("?_host_Info=Outlook%24Win32&spaeter=egal");
  assert.equal(hostInfoParam(), "Outlook$Win32");
});

test("ein angehaengter Anker gehoert nicht mehr zur Kennung", () => {
  // Das Pane setzt selbst Hash-Routen. Liefe der Anker in den Wert hinein, stuende auf
  // der Diagnoseseite eine Kennung, die office.js so nie gesehen hat.
  adresse("?_host_Info=Outlook%24Win32#/diagnose");
  assert.equal(hostInfoParam(), "Outlook$Win32");
});

test("ein Messwert wird ueberschrieben und nicht ein zweites Mal angehaengt", () => {
  // Sonst stuenden nach einem erneuten Versuch zwei widersprechende Zeilen untereinander,
  // und niemand wuesste, welche die aktuelle ist.
  mark("Probe", "erst");
  mark("Probe", "dann");
  const treffer = entries().filter(([name]) => name === "Probe");
  assert.deepEqual(treffer, [["Probe", "dann"]]);
});

test("die Liste der Messwerte ist eine Kopie", () => {
  mark("Fest", "Wert");
  const kopie = entries();
  kopie.push(["Geschmuggelt", "x"]);
  assert.ok(!entries().some(([name]) => name === "Geschmuggelt"));
});

test("die Dauer seit dem Start ist ganzzahlig und laeuft nicht rueckwaerts", () => {
  const a = sinceStart();
  const b = sinceStart();
  assert.equal(Number.isInteger(a), true);
  assert.ok(a >= 0);
  assert.ok(b >= a);
});
