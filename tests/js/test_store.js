/**
 * Der oertliche Speicher tritt an die Stelle der Datenbank, die der Dienst hatte.
 * Geprueft wird, was dabei schiefgehen kann, ohne aufzufallen:
 *
 * - Ein gesperrter Speicher darf das Pane nicht funktionsunfaehig machen.
 * - Ein Techniker darf an einem geteilten Rechner nicht die Vorbelegungen eines
 *   Kollegen sehen, und ein Termin aus einem Testsystem nicht die Kopplung des
 *   Produktivsystems tragen.
 * - Im Speicher steht nie eine Nachrichtenkennung im Klartext.
 * - Der Bestand waechst nicht unbegrenzt.
 */

import test from "node:test";
import assert from "node:assert/strict";

import { Store, digest } from "../../taskpane/js/store.js";

function fakeStorage() {
  const data = new Map();
  return {
    data,
    getItem: (key) => (data.has(key) ? data.get(key) : null),
    setItem: (key, value) => data.set(key, String(value)),
    removeItem: (key) => data.delete(key),
  };
}

function fresh(scope = { instance: "tanss.example", employeeId: 42 }) {
  globalThis.localStorage = fakeStorage();
  return new Store(scope);
}

test("nur bekannte Vorlieben werden uebernommen", () => {
  const store = fresh();
  assert.equal(store.setPref("ticketTypeId", "3"), true);
  assert.equal(store.setPref("lieblingsfarbe", "blau"), false);
  assert.deepEqual(store.prefs(), { ticketTypeId: "3" });
});

test("eine Vorliebe laesst sich wieder loeschen", () => {
  const store = fresh();
  store.setPref("searchScope", "all");
  store.setPref("searchScope", "");
  assert.deepEqual(store.prefs(), {});
});

test("eine kaputte Vorliebe zerlegt die Maske nicht", () => {
  const store = fresh();
  store.setPref("ticketTypeId", "keine Zahl");
  assert.equal(store.intPref("ticketTypeId"), null);
});

test("ein zu langer Wert wird geschnitten, nicht abgewiesen", () => {
  const store = fresh();
  store.setPref("searchScope", "x".repeat(500));
  assert.equal(store.prefs().searchScope.length, 200);
});

test("zwei Mitarbeiter an einem Rechner teilen sich nichts", () => {
  globalThis.localStorage = fakeStorage();
  const anna = new Store({ instance: "t", employeeId: 1 });
  const bert = new Store({ instance: "t", employeeId: 2 });
  anna.setPref("ticketTypeId", "3");
  assert.deepEqual(bert.prefs(), {});
});

test("zwei Installationen teilen sich keine Terminkopplung", () => {
  globalThis.localStorage = fakeStorage();
  const produktiv = new Store({ instance: "produktiv", employeeId: 1 });
  const test_ = new Store({ instance: "test", employeeId: 1 });
  produktiv.rememberAppointmentLink("AAMk123", { supportId: 500 });
  assert.equal(test_.appointmentLink("AAMk123"), null);
});

test("zuletzt benutzte Tickets stehen neueste zuerst und ohne Dublette", () => {
  const store = fresh();
  store.touchTicket(1);
  store.touchTicket(2);
  store.touchTicket(1);
  assert.deepEqual(store.recentTickets(), [1, 2]);
});

test("die Liste der zuletzt benutzten Tickets bleibt kurz", () => {
  const store = fresh();
  for (let i = 1; i <= 25; i += 1) store.touchTicket(i);
  assert.equal(store.recentTickets().length, 10);
  assert.equal(store.recentTickets()[0], 25);
});

test("die Nachrichtenkennung steht nie im Klartext im Speicher", () => {
  const store = fresh();
  const messageId = "<vertraulich-4711@kunde.example>";
  store.rememberAttached(messageId, 77);

  const alles = [...globalThis.localStorage.data.values()].join(" ");
  assert.equal(alles.includes(messageId), false);
  assert.equal(alles.includes("kunde.example"), false);
  assert.equal(alles.includes(digest(messageId)), true);
});

test("die Dublettenwarnung findet dieselbe Nachricht wieder", () => {
  const store = fresh();
  store.rememberAttached("<a@b>", 77);
  store.rememberAttached("<a@b>", 88);
  assert.deepEqual([...store.attachedTicketIds("<a@b>")].sort(), [77, 88]);
  assert.deepEqual([...store.attachedTicketIds("<anderes@b>")], []);
});

test("ohne Nachrichtenkennung wird nichts behauptet", () => {
  const store = fresh();
  assert.deepEqual([...store.attachedTicketIds("")], []);
  assert.equal(store.rememberAttached("", 1), false);
});

test("eine Terminkopplung laesst sich merken und wieder loesen", () => {
  const store = fresh();
  store.rememberAppointmentLink("AAMk123", { supportId: 500, source: "property" });
  assert.equal(store.appointmentLink("AAMk123").supportId, 500);
  store.forgetAppointmentLink("AAMk123");
  assert.equal(store.appointmentLink("AAMk123"), null);
});

test("ein gesperrter Speicher nimmt dem Pane nicht die Funktion", () => {
  globalThis.localStorage = {
    getItem: () => {
      throw new Error("gesperrt");
    },
    setItem: () => {
      throw new Error("gesperrt");
    },
    removeItem: () => {},
  };
  const store = new Store({ instance: "t", employeeId: 1 });
  assert.deepEqual(store.prefs(), {});
  assert.equal(store.setPref("ticketTypeId", "3"), false);
  assert.deepEqual(store.recentTickets(), []);
  assert.deepEqual([...store.attachedTicketIds("<a@b>")], []);
  assert.doesNotThrow(() => store.clear());
});

test("der Bestand der Anhaengevermerke waechst nicht unbegrenzt", () => {
  const store = fresh();
  for (let i = 0; i < 600; i += 1) store.rememberAttached(`<m${i}@b>`, i + 1);
  const records = JSON.parse(globalThis.localStorage.getItem(`${store.prefix}.attached`));
  assert.ok(Object.keys(records).length <= 500, "hoechstens 500 Vermerke");
  // Der juengste muss ueberlebt haben.
  assert.deepEqual([...store.attachedTicketIds("<m599@b>")], [600]);
});
