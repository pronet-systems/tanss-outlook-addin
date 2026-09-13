/**
 * Die Konfiguration kommt aus einer Datei, die ein Mensch auf einem Webserver
 * bearbeitet. Sie muss deshalb drei Dinge aushalten, ohne das Pane zu verschliessen:
 * Sie darf fehlen, sie darf Unsinn enthalten, und sie darf Schluessel tragen, die es
 * nicht gibt.
 *
 * Der Tippfehler ist der wichtigste Fall: Ein `maxEmlByte` ohne `s` darf nicht
 * stillschweigend als neue Einstellung uebernommen werden, denn dann glaubte der
 * Administrator, er habe die Grenze gesetzt, waehrend die Vorgabe gilt.
 */

import test from "node:test";
import assert from "node:assert/strict";

import { DEFAULTS, config, load, note } from "../../taskpane/js/config.js";

/** Legt eine Antwort fuer `fetch("config.json")` bereit. */
function serve(reply) {
  globalThis.fetch = async () => {
    if (reply === null) throw new TypeError("Failed to fetch");
    return new Response(reply.body === undefined ? "" : reply.body, {
      status: reply.status || 200,
      headers: { "Content-Type": "application/json" },
    });
  };
}

test("ohne config.json gelten die Vorgaben und das Pane oeffnet", async () => {
  serve(null);
  const loaded = await load();
  assert.equal(loaded.apiBase, DEFAULTS.apiBase);
  assert.match(note(), /nicht lesbar/);
});

test("eine fehlende Datei ist kein Fehler, sondern eine Vorgabe", async () => {
  serve({ status: 404, body: "nicht da" });
  const loaded = await load();
  assert.equal(loaded.apiBase, DEFAULTS.apiBase);
  assert.match(note(), /HTTP 404/);
});

test("unlesbares JSON schliesst das Pane nicht", async () => {
  serve({ body: "{ das ist kein json" });
  const loaded = await load();
  assert.equal(loaded.apiBase, DEFAULTS.apiBase);
});

test("gesetzte Werte werden uebernommen, auch verschachtelte", async () => {
  serve({
    body: JSON.stringify({
      apiBase: "/tanss/backend",
      frontendBase: "https://tanss.example",
      limits: { maxEmlBytes: 1024 },
      features: { createSupport: true },
      ticketTypes: [{ id: 3, name: "Stoerung" }],
    }),
  });
  const loaded = await load();
  assert.equal(loaded.apiBase, "/tanss/backend");
  assert.equal(loaded.limits.maxEmlBytes, 1024);
  assert.equal(loaded.features.createSupport, true);
  // Nicht genannte Nachbarn behalten ihre Vorgabe.
  assert.equal(loaded.features.appointments, true);
  assert.deepEqual(loaded.ticketTypes, [{ id: 3, name: "Stoerung" }]);
});

test("ein unbekannter Schluessel wird verworfen und benannt", async () => {
  serve({ body: JSON.stringify({ limits: { maxEmlByte: 1 }, apiBse: "/x" }) });
  const loaded = await load();
  assert.equal(loaded.limits.maxEmlBytes, DEFAULTS.limits.maxEmlBytes);
  assert.equal(loaded.apiBase, DEFAULTS.apiBase);
  assert.match(note(), /limits\.maxEmlByte/);
  assert.match(note(), /apiBse/);
});

test("ein falscher Typ wird verworfen, nicht uebernommen", async () => {
  serve({ body: JSON.stringify({ features: { appointments: "ja" } }) });
  const loaded = await load();
  assert.equal(loaded.features.appointments, true);
  assert.match(note(), /features\.appointments/);
});

test("das Kennzeichen kommt aus der TANSS-Adresse, nicht aus der Seitenadresse", async () => {
  // Der Unterschied faellt erst bei einer gemeinsam genutzten Ablage auf, dann aber
  // hart: Dort ist die Seitenadresse fuer ALLE Kunden dieselbe. Zwei Instanzen auf
  // einem Arbeitsplatz teilten sich sonst ihren oertlichen Bereich, und ein Termin
  // truege die Einsatznummer der jeweils anderen Installation.
  globalThis.location = { href: "https://gemeinsam.example/", hostname: "gemeinsam.example" };
  serve({ body: JSON.stringify({ apiBase: "https://tanss.kunde-a.de/backend" }) });
  assert.equal((await load()).instance, "tanss.kunde-a.de");

  serve({ body: JSON.stringify({ apiBase: "https://tanss.kunde-b.de/backend" }) });
  assert.equal((await load()).instance, "tanss.kunde-b.de");
  delete globalThis.location;
});

test("liegt die API im selben Ursprung, traegt die Seitenadresse das Kennzeichen", async () => {
  globalThis.location = { href: "https://tanss.example.de/", hostname: "tanss.example.de" };
  serve({ body: JSON.stringify({ apiBase: "/backend" }) });
  assert.equal((await load()).instance, "tanss.example.de");
  delete globalThis.location;
});

test("ein ausdruecklich gesetztes Kennzeichen schlaegt jede Ableitung", async () => {
  globalThis.location = { href: "https://gemeinsam.example/", hostname: "gemeinsam.example" };
  serve({ body: JSON.stringify({ apiBase: "https://tanss.kunde.de/backend", instance: "produktiv" }) });
  assert.equal((await load()).instance, "produktiv");
  delete globalThis.location;
});

test("config() liefert nach dem Laden dasselbe wie load()", async () => {
  serve({ body: JSON.stringify({ apiBase: "/a" }) });
  const loaded = await load();
  assert.equal(config(), loaded);
});

/* --------------------------------------------------- Welche TANSS-Instanz gilt */

/**
 * Ein gemeinsamer Hoster bedient beliebig viele Kunden, weil jeder sein eigenes
 * Manifest mitbringt und darin seine TANSS-Adresse steht. Das ist bequem und zugleich
 * die empfindlichste Stelle des ganzen Aufbaus: Der Techniker tippt sein Kennwort in
 * dieses Pane. Eine untergeschobene Adresse waere eine Anmeldemaske, die es woanders
 * hin schickt.
 */

import { apiHost, isAllowed, setApiBase } from "../../taskpane/js/config.js";

test("die Adresse aus dem Manifest schlaegt die aus der Datei", async () => {
  globalThis.location = { href: "https://addon.example/?tanss=https://kunde.example/backend" };
  serve({ body: JSON.stringify({ apiBase: "https://andere.example/backend" }) });
  const loaded = await load();
  assert.equal(loaded.apiBase, "https://kunde.example/backend");
  delete globalThis.location;
});

test("ohne Angabe in der Adresse gilt die Datei", async () => {
  globalThis.location = { href: "https://addon.example/" };
  serve({ body: JSON.stringify({ apiBase: "https://tanss.example/backend" }) });
  const loaded = await load();
  assert.equal(loaded.apiBase, "https://tanss.example/backend");
  delete globalThis.location;
});

test("http wird niemals zugelassen - das Kennwort liefe im Klartext", () => {
  assert.equal(isAllowed("http://tanss.example/backend", []), false);
  assert.equal(isAllowed("https://tanss.example/backend", []), true);
});

test("eine gefuellte Liste weist fremde Adressen ab", async () => {
  globalThis.location = { href: "https://addon.example/?tanss=https://boese.example/backend" };
  serve({
    body: JSON.stringify({
      apiBase: "https://tanss.example/backend",
      allowedApiOrigins: ["https://tanss.example"],
    }),
  });
  const loaded = await load();
  assert.equal(loaded.apiBase, "https://tanss.example/backend",
    "die untergeschobene Adresse darf nicht gewinnen");
  delete globalThis.location;
});

test("eine selbst eingetragene Adresse wird gemerkt und geprueft", async () => {
  globalThis.localStorage = {
    store: new Map(),
    getItem(key) { return this.store.has(key) ? this.store.get(key) : null; },
    setItem(key, value) { this.store.set(key, value); },
    removeItem(key) { this.store.delete(key); },
  };
  globalThis.location = { href: "https://addon.example/" };
  serve({ body: JSON.stringify({ apiBase: "" }) });
  await load();

  assert.equal(setApiBase("http://unsicher.example"), false);
  assert.equal(setApiBase("https://meine.example/backend"), true);

  const wieder = await load();
  assert.equal(wieder.apiBase, "https://meine.example/backend");
  delete globalThis.location;
  delete globalThis.localStorage;
});

test("der angezeigte Name stammt aus der wirklich benutzten Adresse", async () => {
  globalThis.location = { href: "https://addon.example/?tanss=https://tanss.kunde.de/backend" };
  serve({ body: "{}" });
  await load();
  assert.equal(apiHost(), "tanss.kunde.de");
  delete globalThis.location;
});
