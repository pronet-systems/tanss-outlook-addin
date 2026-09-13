/**
 * Die Terminkennung. Zwei Regeln, und beide brechen lautlos:
 *
 * 1. Outlook bettet die eigentliche UID in eine laengere Kette ein. Wer die Kette
 *    selbst vergleicht, findet eine bestehende Kopplung nicht wieder - dieselbe
 *    Besprechung traegt in zwei Postfaechern verschiedene Ketten, aber dieselbe
 *    eingebettete UID.
 * 2. Der Server vergleicht ab einer gewissen Laenge nur Anfang und Ende. Wer strenger
 *    prueft als er, uebersieht Kopplungen, die es gibt - die Aufloesung faellt dann
 *    still auf den Zeitfenster-Abgleich zurueck und trifft eine Faustregel statt einer
 *    Tatsache.
 */

import test from "node:test";
import assert from "node:assert/strict";

import {
  NO_SEQUENCE,
  canonicalUid,
  formatSyncGroup,
  isAmbiguouslyLong,
  parseSyncGroup,
  syncGroupMatches,
  uidMatches,
} from "../../taskpane/js/tanss/uid.js";

/** Baut eine Kette, wie Outlook sie liefert: Vorspann, Marke, Laengenfeld, UID, Fuellung. */
function globalObjectId(innerUid) {
  const bytes = [];
  for (const ch of "040000008200E00074C5B7101A82E008") bytes.push(ch.charCodeAt(0));
  for (const ch of "vCal-Uid") bytes.push(ch.charCodeAt(0));
  bytes.push(1, 0, 0, 0);
  for (const ch of innerUid) bytes.push(ch.charCodeAt(0));
  bytes.push(0, 0, 0);
  return bytes.map((b) => b.toString(16).padStart(2, "0")).join("");
}

test("die eingebettete Kennung wird herausgeloest", () => {
  const inner = "040000008200E00074C5B7101A82E00807D6@example.de";
  assert.equal(canonicalUid(globalObjectId(inner)), inner);
});

test("dieselbe Besprechung in zwei Postfaechern ergibt dieselbe Kennung", () => {
  // Die Ketten unterscheiden sich in der Fuellung, die eingebettete UID nicht.
  const inner = "besprechung-4711@example.de";
  const a = globalObjectId(inner);
  const b = `${globalObjectId(inner)}0000`;
  assert.equal(canonicalUid(a), canonicalUid(b));
});

test("was keine Kette ist, bleibt unveraendert", () => {
  assert.equal(canonicalUid("schlichte-uid@example.de"), "schlichte-uid@example.de");
  // Ungerade Laenge, also keine Hex-Kette.
  assert.equal(canonicalUid("abc"), "abc");
  assert.equal(canonicalUid(""), "");
  assert.equal(canonicalUid(null), "");
});

test("eine Hex-Kette ohne Marke bleibt ebenfalls unveraendert", () => {
  const ohneMarke = "00112233445566778899aabbccddeeff00112233445566778899aabbccddeeff";
  assert.equal(canonicalUid(ohneMarke), ohneMarke);
});

test("kurze Kennungen werden streng verglichen", () => {
  assert.equal(uidMatches("abc@x", "abc@x"), true);
  assert.equal(uidMatches("abc@x", "abd@x"), false);
  assert.equal(uidMatches("", "abc@x"), false);
  assert.equal(uidMatches(null, null), false);
});

test("lange Kennungen werden wie beim Server nur an den Enden verglichen", () => {
  const kopf = "K".repeat(32);
  const fuss = "F".repeat(70);
  const a = `${kopf}${"M".repeat(20)}${fuss}`;
  const b = `${kopf}${"N".repeat(20)}${fuss}`;
  assert.ok(a.length >= 100 && b.length >= 100);
  assert.equal(uidMatches(a, b), true, "der Mittelteil darf nicht entscheiden");

  const anderesEnde = `${kopf}${"M".repeat(20)}${"G".repeat(70)}`;
  assert.equal(uidMatches(a, anderesEnde), false);
});

test("eine Kennung ueber der Grenze wird als mehrdeutig gemeldet", () => {
  assert.equal(isAmbiguouslyLong("x".repeat(99)), false);
  assert.equal(isAmbiguouslyLong("x".repeat(100)), true);
  assert.equal(isAmbiguouslyLong(""), false);
});

test("ist eine Seite kurz, gilt strikte Gleichheit trotz Laenge der anderen", () => {
  const lang = "x".repeat(120);
  assert.equal(uidMatches(lang, "x".repeat(99)), false);
});

test("der Kopplungswert wird in Kennung und Wiederholung zerlegt", () => {
  assert.deepEqual(parseSyncGroup("uid@x|7"), { uid: "uid@x", sequence: 7 });
  assert.deepEqual(parseSyncGroup("uid@x"), { uid: "uid@x", sequence: NO_SEQUENCE });
  assert.deepEqual(parseSyncGroup(""), { uid: "", sequence: NO_SEQUENCE });
});

test("die Wiederholung null wird nicht als 'nicht gesetzt' verschluckt", () => {
  assert.deepEqual(parseSyncGroup("uid@x|0"), { uid: "uid@x", sequence: 0 });
});

test("ein Trennzeichen, das keine Wiederholung ist, gehoert zur Kennung", () => {
  assert.deepEqual(parseSyncGroup("uid|mit|text"),
    { uid: "uid|mit|text", sequence: NO_SEQUENCE });
});

test("der Kopplungswert wird ohne Wiederholung ohne Zusatz gebaut", () => {
  assert.equal(formatSyncGroup("uid@x"), "uid@x");
  assert.equal(formatSyncGroup("uid@x", NO_SEQUENCE), "uid@x");
  assert.equal(formatSyncGroup("uid@x", 3), "uid@x|3");
});

test("ein gelesener Wert ohne Wiederholung passt zur Kennung", () => {
  // Der Regelfall beim Lesen: expandierte Wiederholungen erben die Angaben des
  // Ursprungstermins und tragen selbst keine Wiederholungsnummer.
  assert.equal(syncGroupMatches("uid@x", "uid@x", 5), true);
  assert.equal(syncGroupMatches("uid@x|5", "uid@x", 5), true);
  assert.equal(syncGroupMatches("uid@x|4", "uid@x", 5), false);
  assert.equal(syncGroupMatches("andere@x", "uid@x"), false);
  assert.equal(syncGroupMatches("", "uid@x"), false);
});
