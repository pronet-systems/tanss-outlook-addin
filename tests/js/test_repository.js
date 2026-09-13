/**
 * Die Fachschicht. Geprueft wird, was im Betrieb nicht auffaellt, wenn es bricht:
 *
 * - Eine leere Trefferliste ist etwas anderes als eine gekappte. "Leer" heisst "nichts
 *   gefunden", "leer plus Hinweis" heisst "Suche verfeinern".
 * - Namen kommen aus dem meta-Block derselben Antwort, nicht aus Zusatzaufrufen.
 * - Eine Ziffernfolge ist eine Ticketnummer und keine Suche.
 * - Ausgeschiedene Ansprechpartner werden nicht als Melder angeboten.
 * - Ein Erfolg ohne Beleg ist kein Erfolg.
 * - "Nicht nachgesehen" ist nicht "nichts gefunden".
 */

import test from "node:test";
import assert from "node:assert/strict";

import { TanssRepository, isFreshEmlDocument } from "../../taskpane/js/tanss/repository.js";

/** Ein Client, der Aufrufe aufzeichnet und vorbereitete Antworten liefert. */
function fakeClient(routes) {
  const calls = [];
  const answer = (method, path, body, options) => {
    calls.push({ method, path, body, options });
    const key = `${method} ${path}`;
    if (!(key in routes)) throw new Error(`ungeplante Route: ${key}`);
    const reply = routes[key];
    if (reply instanceof Error) throw reply;
    return reply;
  };
  return {
    calls,
    get: async (path, options = {}) => {
      const result = answer("GET", path, null, options);
      return options.wantMeta ? result : (result.content ?? result);
    },
    put: async (path, json, options = {}) => {
      const result = answer("PUT", path, json, options);
      return options.wantMeta ? result : (result.content ?? result);
    },
    post: async (path, json, options = {}) => {
      const result = answer("POST", path, json, options);
      return options.wantMeta ? result : (result.content ?? result);
    },
    call: async (method, path, options = {}) => {
      const result = answer(method, path, options.body, options);
      return result.content ?? result;
    },
  };
}

const repo = (routes, config = {}) =>
  new TanssRepository({ client: fakeClient(routes), config });

/* ------------------------------------------------------------------- Suche */

test("Firmennamen kommen aus dem meta-Block, ohne Zusatzaufruf", async () => {
  const client = fakeClient({
    "PUT /api/v1/search": {
      content: { tickets: [{ id: 7, title: "Drucker", companyId: 3, statusId: 9 }] },
      meta: {
        linkedEntities: {
          companies: { 3: { name: "Muster GmbH" } },
          ticketStates: { 9: { name: "In Arbeit" } },
        },
      },
    },
  });
  const repository = new TanssRepository({ client, config: {} });
  const { items } = await repository.searchTickets("Drucker");

  assert.equal(items[0].companyName, "Muster GmbH");
  assert.equal(items[0].statusName, "In Arbeit");
  assert.equal(client.calls.length, 1, "kein Zusatzaufruf je Treffer");
});

test("leer plus Hinweis heisst 'Suche verfeinern', nicht 'nichts gefunden'", async () => {
  const gekappt = await repo({
    "PUT /api/v1/search": {
      content: { companies: [] },
      meta: { properties: { extras: { tooManyEntries: true } } },
    },
  }).searchCompanies("a");
  assert.deepEqual(gekappt.items, []);
  assert.equal(gekappt.tooMany, true);

  const leer = await repo({
    "PUT /api/v1/search": { content: { companies: [] }, meta: {} },
  }).searchCompanies("zzz");
  assert.equal(leer.tooMany, false);
});

test("der Hinweis wird auch im zweiten Behaelter gefunden", async () => {
  const { tooMany } = await repo({
    "PUT /api/v1/search": {
      content: { companies: [] },
      meta: { extras: { tooManyEntries: true } },
    },
  }).searchCompanies("a");
  assert.equal(tooMany, true);
});

test("ausgeschiedene Ansprechpartner werden nicht als Melder angeboten", async () => {
  const client = fakeClient({ "PUT /api/v1/search": { content: { employees: [] }, meta: {} } });
  await new TanssRepository({ client, config: {} }).searchEmployees("muster", { companyId: 3 });
  assert.equal(client.calls[0].body.configs.employee.inactive, false);
});

test("die Vorschaulaenge wird immer gesetzt", async () => {
  const client = fakeClient({ "PUT /api/v1/search": { content: { tickets: [] }, meta: {} } });
  await new TanssRepository({ client, config: {} }).searchTickets("drucker");
  assert.equal(client.calls[0].body.configs.ticket.previewContentMaxChars, 120);
});

test("eine Ziffernfolge ist eine Ticketnummer und keine Suche", async () => {
  const client = fakeClient({
    "PUT /api/v1/tickets": { content: [{ id: 4711, title: "Treffer" }], meta: {} },
  });
  const { items } = await new TanssRepository({ client, config: {} }).searchTickets("4711");
  assert.equal(client.calls[0].path, "/api/v1/tickets");
  assert.deepEqual(client.calls[0].body, { ids: [4711] });
  assert.equal(items[0].id, 4711);
});

test("eine leere Suche ohne Firma fragt gar nicht erst", async () => {
  const client = fakeClient({});
  const { items } = await new TanssRepository({ client, config: {} }).searchTickets("");
  assert.deepEqual(items, []);
  assert.equal(client.calls.length, 0);
});

test("bereits angehaengte Mails werden am Treffer gekennzeichnet", async () => {
  const { items } = await repo({
    "PUT /api/v1/tickets": { content: [{ id: 1 }, { id: 2 }], meta: {} },
  }).searchTickets("", { companyId: 3, attachedTicketIds: new Set([2]) });
  assert.equal(items[0].alreadyHasThisMail, false);
  assert.equal(items[1].alreadyHasThisMail, true);
});

/* ---------------------------------------------------------------- Ticketmaske */

test("faellt die Feldsteuerung aus, gilt die mildere Annahme", async () => {
  const options = await repo({
    "GET /api/v1/tickets/": new Error("weg"),
    "GET /api/v1/admin/ticketStates": { content: [] },
    "GET /api/v1/employees/technicians": { content: [] },
    "GET /api/v1/employees/departments": { content: [] },
  }).ticketOptions();
  assert.equal(options.remitterRequired, false);
  assert.equal(options.forceAssignment, false);
  assert.equal(options.source, "fallback");
});

test("eine Antwort ohne Feldsteuerung wird nicht als Feldsteuerung gelesen", async () => {
  const options = await repo({
    // Die Route antwortete wie eine gewoehnliche Ticketliste.
    "GET /api/v1/tickets/": { content: [{ id: 1 }], meta: {} },
    "GET /api/v1/admin/ticketStates": { content: [] },
    "GET /api/v1/employees/technicians": { content: [] },
    "GET /api/v1/employees/departments": { content: [] },
  }).ticketOptions();
  assert.equal(options.source, "fallback");
});

test("die Feldregeln der Instanz werden uebernommen, wenn es sie gibt", async () => {
  const options = await repo({
    "GET /api/v1/tickets/": {
      content: {},
      meta: {
        properties: {
          extras: { remitterIsAMandatoryField: true, forceAssignment: true,
            autoAssignedEmployeeId: 12 },
        },
      },
    },
    "GET /api/v1/admin/ticketStates": { content: [] },
    "GET /api/v1/employees/technicians": { content: [] },
    "GET /api/v1/employees/departments": { content: [] },
  }).ticketOptions();
  assert.equal(options.remitterRequired, true);
  assert.equal(options.autoAssignedEmployeeId, 12);
  assert.equal(options.source, "properties");
});

test("Status kommen sortiert und nur aktiv", async () => {
  const options = await repo({
    "GET /api/v1/tickets/": new Error("weg"),
    "GET /api/v1/admin/ticketStates": {
      content: [
        { id: 2, name: "Zweiter", rank: 2, active: true },
        { id: 9, name: "Stillgelegt", rank: 0, active: false },
        { id: 1, name: "Erster", rank: 1, active: true },
      ],
    },
    "GET /api/v1/employees/technicians": { content: [] },
    "GET /api/v1/employees/departments": { content: [] },
  }).ticketOptions();
  assert.deepEqual(options.states.map((s) => s.id), [1, 2]);
});

test("eine ausgefallene Auswahlliste bleibt leer und reisst nichts mit", async () => {
  const options = await repo({
    "GET /api/v1/tickets/": new Error("weg"),
    "GET /api/v1/admin/ticketStates": new Error("weg"),
    "GET /api/v1/employees/technicians": { content: [{ id: 5, name: "Anna" }] },
    "GET /api/v1/employees/departments": new Error("weg"),
  }).ticketOptions();
  assert.deepEqual(options.states, []);
  assert.deepEqual(options.departments, []);
  assert.equal(options.technicians[0].name, "Anna");
});

test("Leistungsarten fallen auf die konfigurierte Vorgabe zurueck", async () => {
  const types = await repo(
    { "GET /api/v1/supportTypes/active": new Error("weg") },
    { defaults: { supportTypeId: 4 } },
  ).supportTypes();
  assert.deepEqual(types, [{ id: 4, name: "Vorgabe" }]);
});

/* ------------------------------------------------------------------ Anlegen */

test("der Titel wird sichtbar auf 100 Zeichen geschnitten", async () => {
  const client = fakeClient({
    "POST /api/v1/tickets": { content: { id: 1 }, meta: {} },
  });
  await new TanssRepository({ client, config: {} })
    .createTicket({ title: "x".repeat(250), content: "" });
  assert.equal(client.calls[0].body.title.length, 100);
});

test("eine unvollstaendige Zuordnung wird gar nicht erst gesendet", async () => {
  const client = fakeClient({ "POST /api/v1/tickets": { content: { id: 1 }, meta: {} } });
  await new TanssRepository({ client, config: {} })
    .createTicket({ title: "T", linkId: 5 });
  assert.equal("linkId" in client.calls[0].body, false);
  assert.equal("linkTypeId" in client.calls[0].body, false);
});

test("ein Erfolg ohne Ticketnummer gilt als Fehlschlag", async () => {
  await assert.rejects(
    () => repo({ "POST /api/v1/tickets": { content: {}, meta: {} } })
      .createTicket({ title: "T" }),
    (error) => error.code === "INTERNAL",
  );
});

test("ein Ansprechpartner wird am letzten Leerzeichen getrennt", async () => {
  const client = fakeClient({
    "POST /api/v1/employees": { content: { id: 8, name: "Anna Maria Muster" } },
  });
  await new TanssRepository({ client, config: {} })
    .createContact({ companyId: 3, name: "Anna Maria Muster", email: "a@example.de" });
  assert.equal(client.calls[0].body.firstName, "Anna Maria");
  assert.equal(client.calls[0].body.lastName, "Muster");
  assert.deepEqual(client.calls[0].body.companyAssignments, [{ companyId: 3 }]);
});

test("eine Antwort als Liste wird genauso gelesen wie ein einzelnes Objekt", async () => {
  const contact = await repo({
    "POST /api/v1/employees": { content: [{ id: 8, name: "Muster" }] },
  }).createContact({ companyId: 3, name: "Muster", email: "a@example.de" });
  assert.equal(contact.id, 8);
});

/* ---------------------------------------------------------------- Mailablage */

test("eine zu grosse Datei wird nicht als Erfolg gemeldet", async () => {
  const result = await repo({
    "GET /api/v1/tickets/history/5": { content: { mails: [] } },
    "POST /api/v1/tickets/5/upload": { content: [{ status: "TOO_BIG", name: "a.eml", size: 9 }] },
  }).uploadEml(5, { filename: "a.eml", blob: new Blob(["x"]), description: "d" });
  assert.equal(result.status, "TOO_BIG");
  assert.equal(result.verified, false);
});

test("ein Upload ohne Ergebnisliste gilt als Fehlschlag", async () => {
  await assert.rejects(
    () => repo({
      "GET /api/v1/tickets/history/5": { content: { mails: [] } },
      "POST /api/v1/tickets/5/upload": { content: [] },
    }).uploadEml(5, { filename: "a.eml", blob: new Blob(["x"]), description: "d" }),
    (error) => error.code === "INTERNAL",
  );
});

test("die neue Nachricht wird ueber die Differenz erkannt, nicht ueber den Betreff", async () => {
  let round = 0;
  const client = {
    calls: [],
    get: async (path) => {
      if (path === "/api/v1/tickets/history/5") {
        round += 1;
        return round === 1 ? { mails: [{ id: 1 }] } : { mails: [{ id: 1 }, { id: 2 }] };
      }
      return [];
    },
    put: async () => ({}),
    post: async () => ({}),
    call: async () => [{ status: "OK", name: "a.eml", size: 9 }],
  };
  const result = await new TanssRepository({ client, config: {} })
    .uploadEml(5, { filename: "a.eml", blob: new Blob(["x"]), description: "d" });
  assert.equal(result.tanssMailId, 2);
  assert.equal(result.verified, true);
});

test("'nicht nachgesehen' ist nicht 'nichts gefunden'", async () => {
  const client = {
    get: async (path) => {
      if (path.startsWith("/api/v1/tickets/history/")) throw new Error("weg");
      return [];
    },
    put: async () => ({}),
    post: async () => ({}),
    call: async () => [{ status: "OK", name: "a.eml", size: 9 }],
  };
  const result = await new TanssRepository({ client, config: {} })
    .uploadEml(5, { filename: "a.eml", blob: new Blob(["x"]), description: "d" });
  // Ohne Vergleichsstand darf keine Kennung behauptet werden.
  assert.equal(result.tanssMailId, null);
});

test("mehrere neue Nachrichten gleichzeitig ergeben keine Kennung", async () => {
  let round = 0;
  const client = {
    get: async (path) => {
      if (path === "/api/v1/tickets/history/5") {
        round += 1;
        return round === 1 ? { mails: [] } : { mails: [{ id: 2 }, { id: 3 }] };
      }
      return [];
    },
    put: async () => ({}),
    post: async () => ({}),
    call: async () => [{ status: "OK", name: "a.eml", size: 9 }],
  };
  const result = await new TanssRepository({ client, config: {} })
    .uploadEml(5, { filename: "a.eml", blob: new Blob(["x"]), description: "d" });
  assert.equal(result.tanssMailId, null);
  assert.equal(result.verified, true, "geprueft wurde trotzdem");
});

test("eine still als Dokument abgelegte Nachricht wird erkannt", () => {
  const now = 1_700_000_000;
  assert.equal(
    isFreshEmlDocument({ fileName: `eml_5_${now}_post.eml`, date: 0 }, 5, now),
    true,
  );
  // Anderes Ticket.
  assert.equal(
    isFreshEmlDocument({ fileName: `eml_9_${now}_post.eml`, date: now }, 5, now),
    false,
  );
  // Zu alt - das ist eine fruehere Ablage, nicht die eigene.
  assert.equal(
    isFreshEmlDocument({ fileName: `eml_5_${now - 3600}_post.eml`, date: now - 3600 }, 5, now),
    false,
  );
  // Ein gewoehnliches Dokument.
  assert.equal(isFreshEmlDocument({ fileName: "Angebot.pdf", date: now }, 5, now), false);
});
