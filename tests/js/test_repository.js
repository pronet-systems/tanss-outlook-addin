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

import { TanssRepository, departmentIdsOf, isFreshEmlDocument, shapeOf } from "../../taskpane/js/tanss/repository.js";
import { ApiError, reasonOf } from "../../taskpane/js/tanss/errors.js";

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

test("die Feldregeln sind eine Annahme, kein Aufruf", async () => {
  // Frueher wurde dafuer `GET /api/v1/tickets/` gerufen. Diese Route gibt es nicht -
  // unter /api/v1/tickets sind nur POST und PUT definiert. Der Aufruf kostete bei jedem
  // Oeffnen der Maske eine Umlaufzeit und landete zuverlaessig im Rueckfall; die Maske
  // meldete deshalb IMMER, die Listen staemmten aus dem Rueckfall. Das war kein Ausfall,
  // sondern ein Entwurfsfehler - und die Meldung schob ihn einer Gegenstelle zu.
  const client = fakeClient({
    "GET /api/v1/admin/ticketTypes": { content: [] },
    "GET /api/v1/admin/ticketStates": { content: [] },
    "GET /api/v1/employees/technicians": { content: [] },
    "GET /api/v1/employees/departments": { content: [] },
  });
  const options = await new TanssRepository({ client, config: {} }).ticketOptions();

  assert.equal(options.remitterRequired, false);
  assert.equal(options.forceAssignment, false);
  assert.equal("autoAssignedEmployeeId" in options, false,
    "das Feld kommt in TANSS nicht vor - es darf keine Vorbelegung erzeugen");
  assert.equal(client.calls.some((c) => c.path === "/api/v1/tickets/"), false,
    "die Route gibt es nicht; sie darf nicht mehr gerufen werden");
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

test("ohne Ergebnisliste UND ohne Nachricht am Ticket ist der Fehlschlag belegt", async () => {
  // Frueher wurde hier geworfen, WEIL die Liste fehlte. Das war ein Schluss aus dem
  // Schweigen. Jetzt wird nachgesehen, und erst der Befund am Ticket entscheidet - dann
  // ist der Fehlschlag belegt statt vermutet, und die Wiederholung ist gefahrlos.
  await assert.rejects(
    () => repo({
      "GET /api/v1/tickets/history/5": { content: { mails: [] } },
      "POST /api/v1/tickets/5/upload": { content: [] },
    }).uploadEml(5, { filename: "a.eml", blob: new Blob(["x"]), description: "d" }),
    (error) => error.code === "INTERNAL" && /haengt die Nachricht nicht/.test(error.message),
  );
});

test("ohne Ergebnisliste, aber die Nachricht haengt am Ticket: Erfolg aus der Gegenprobe", async () => {
  // Der Fall, der im ersten echten Lauf eingetreten ist: TANSS bestaetigte den Upload
  // ohne Ergebniszeile, und die Mail hing trotzdem am Ticket. Frueher meldete das Pane
  // hier einen Fehlschlag und bot die Wiederholung an - ein Klick darauf haette die
  // Nachricht ein ZWEITES Mal angehaengt.
  let runde = 0;
  const client = {
    calls: [],
    get: async (pfad) => {
      if (pfad === "/api/v1/tickets/history/5") {
        runde += 1;
        return runde === 1 ? { mails: [] } : { mails: [{ id: 77 }] };
      }
      return [];
    },
    put: async () => ({}),
    post: async () => ({}),
    call: async () => [],
  };
  const ergebnis = await new TanssRepository({ client, config: {} })
    .uploadEml(5, { filename: "a.eml", blob: new Blob(["x"]), description: "d" });

  assert.equal(ergebnis.status, "OK");
  assert.equal(ergebnis.confirmedBy, "check", "die Herkunft der Bestaetigung muss mitkommen");
  assert.equal(ergebnis.tanssMailId, 77);
  assert.equal(ergebnis.verified, true);
});

test("ohne Ergebnisliste und ohne Gegenprobe wird nichts behauptet", async () => {
  // Kommt auch die Gegenprobe nicht durch, ist NICHTS feststellbar. Dann darf weder ein
  // Erfolg noch ein belegter Fehlschlag gemeldet werden - die Meldung schickt den
  // Techniker nach TANSS, bevor er es noch einmal versucht.
  const client = {
    calls: [],
    get: async () => { throw new Error("weg"); },
    put: async () => ({}),
    post: async () => ({}),
    call: async () => [],
  };
  await assert.rejects(
    () => new TanssRepository({ client, config: {} })
      .uploadEml(5, { filename: "a.eml", blob: new Blob(["x"]), description: "d" }),
    (error) => error.code === "INTERNAL" && /nachsehen/.test(error.message),
  );
});

test("ein einzelnes Ergebnisobjekt zaehlt wie eine Liste mit einem Eintrag", async () => {
  // Dieselbe Nachsicht, die `createContact` an derselben Schnittstelle schon kennt: als
  // Liste beschrieben, in der Praxis auch ein einzelnes Objekt.
  const client = {
    calls: [],
    get: async () => ({ mails: [] }),
    put: async () => ({}),
    post: async () => ({}),
    call: async () => ({ status: "TOO_BIG", name: "a.eml", size: 9 }),
  };
  const ergebnis = await new TanssRepository({ client, config: {} })
    .uploadEml(5, { filename: "a.eml", blob: new Blob(["x"]), description: "d" });

  assert.equal(ergebnis.status, "TOO_BIG");
});

test("die Meldung nennt die Gestalt der Antwort, nie ihren Inhalt", async () => {
  // "ohne Ergebnisliste" liess niemanden wissen, was TANSS statt dessen geschickt hat.
  // Genannt werden Typ und Feldnamen - Feldnamen beschreiben die Schnittstelle, Werte
  // beschrieben den Vorgang, und der geht niemanden ausser dem Techniker etwas an.
  assert.equal(shapeOf(null), "leere Antwort");
  assert.equal(shapeOf([]), "leere Liste (0 Einträge)");
  assert.equal(shapeOf("nichts"), "string statt Liste");
  assert.equal(shapeOf({}), "leeres Objekt statt Liste");
  assert.equal(shapeOf({ ok: true, hinweis: "x" }), "Objekt statt Liste, Felder: ok, hinweis");
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

/* ------------------------------------------------------------- Firmenbindung */

test("ein Mitarbeiter einer fremden Firma wird gar nicht erst geliefert", async () => {
  // Der gemeldete Fehler aus dem ersten echten Lauf: Es liess sich ein Ansprechpartner
  // waehlen, der nicht zur gewaehlten Firma gehoert. Die Anfrage gibt die Firma zwar mit,
  // aber ob TANSS danach filtert, ist eine Zusage - und eine ungepruefte Zusage ist keine.
  // Ein Ticket am falschen Kunden mit dem Namen einer fremden Person ist kein
  // Schoenheitsfehler, deshalb steht die Sperre hier und nicht in der Maske.
  const client = fakeClient({
    "PUT /api/v1/search": {
      content: {
        employees: [
          { id: 1, name: "Eigene", companyAssignments: [{ companyId: 3 }] },
          { id: 2, name: "Fremde", companyAssignments: [{ companyId: 9 }] },
        ],
      },
      meta: {},
    },
  });
  const ergebnis = await new TanssRepository({ client, config: {} })
    .searchEmployees("mu", { companyId: 3 });

  assert.deepEqual(ergebnis.items.map((r) => r.name), ["Eigene"]);
  assert.equal(ergebnis.unverified, false);
});

test("ein Mitarbeiter mit mehreren Firmen bleibt, wenn die gesuchte dabei ist", async () => {
  // In TANSS ist die Zuordnung eine MENGE - `employeeWrite` schreibt sie als
  // `companyAssignments`. Wer hier ein einzelnes Feld erwartete, hielte einen
  // Ansprechpartner mit zwei Firmen bei einer davon faelschlich fuer fremd und liesse den
  // Techniker einen Melder nicht waehlen, der voellig in Ordnung ist.
  const client = fakeClient({
    "PUT /api/v1/search": {
      content: {
        employees: [{ id: 1, name: "Doppelt", companyAssignments: [{ companyId: 9 }, { companyId: 3 }] }],
      },
      meta: {},
    },
  });
  const ergebnis = await new TanssRepository({ client, config: {} })
    .searchEmployees("mu", { companyId: 3 });

  assert.deepEqual(ergebnis.items.map((r) => r.name), ["Doppelt"]);
  assert.equal(ergebnis.unverified, false);
});

test("nennt TANSS keine Firma, wird die Zeile gezeigt und die Luecke benannt", async () => {
  // Der dritte Ausgang, und der wichtigste: "nicht feststellbar" ist weder "gehoert dazu"
  // noch "gehoert nicht dazu". Wer ihn als "gehoert nicht" behandelte, leerte die
  // Trefferliste jeder Instanz, die das Feld nicht mitliefert - aus einer
  // Schutzmassnahme wuerde ein Ausfall. Wer ihn stillschweigend durchwinkt, behauptet
  // eine Pruefung, die nicht stattgefunden hat.
  const client = fakeClient({
    "PUT /api/v1/search": {
      content: { employees: [{ id: 1, name: "Ohne Angabe" }] },
      meta: {},
    },
  });
  const ergebnis = await new TanssRepository({ client, config: {} })
    .searchEmployees("mu", { companyId: 3 });

  assert.deepEqual(ergebnis.items.map((r) => r.name), ["Ohne Angabe"]);
  assert.equal(ergebnis.unverified, true, "die Luecke muss gemeldet werden");
});

test("ohne gefragte Firma wird nichts gefiltert und nichts behauptet", async () => {
  // Die Suche ueber alle Kunden gibt es weiterhin - die Ticketsuche benutzt sie. Sie darf
  // aber keine Firmenpruefung vortaeuschen, die niemand verlangt hat.
  const client = fakeClient({
    "PUT /api/v1/search": {
      content: { employees: [{ id: 2, name: "Fremde", companyAssignments: [{ companyId: 9 }] }] },
      meta: {},
    },
  });
  const ergebnis = await new TanssRepository({ client, config: {} }).searchEmployees("mu");

  assert.deepEqual(ergebnis.items.map((r) => r.name), ["Fremde"]);
  assert.equal(ergebnis.unverified, false);
});

/* -------------------------------------------------------- Gruende statt Schweigen */

test("faellt eine Auswahlliste aus, steht der Grund in der Antwort", async () => {
  // Zweimal hintereinander hat das Pane gemeldet, die Technikerliste sei "nicht
  // abrufbar" - und niemand konnte sagen, woran es liegt, weil der Grund zwei Zeilen
  // vorher in einem leeren `catch` verschwand. Ein Ausfall darf entbehrlich sein; er
  // darf nicht unbemerkt sein.
  const client = {
    calls: [],
    get: async (pfad) => {
      if (pfad === "/api/v1/employees/technicians") {
        throw new ApiError("INTERNAL", "weg", { status: 404 });
      }
      return [];
    },
    put: async () => ({ content: {}, meta: {} }),
    post: async () => ({}),
    call: async () => [],
  };
  const options = await new TanssRepository({ client, config: {} }).ticketOptions({});

  const treffer = options.failures.find((f) => f.path === "/api/v1/employees/technicians");
  assert.ok(treffer, `kein Eintrag fuer die Technikerliste: ${JSON.stringify(options.failures)}`);
  assert.match(treffer.reason, /404/, "der HTTP-Status muss dastehen");
  assert.deepEqual(options.technicians, [], "die Liste bleibt trotzdem leer statt zu werfen");
});

test("der Grund nennt die Schnittstelle, nie den Antwortkoerper", async () => {
  // Er wandert in eine Meldung, die weitergereicht wird. Route, Code und Status
  // beschreiben die Schnittstelle; ein Antwortkoerper beschriebe den Vorgang.
  assert.equal(reasonOf(new ApiError("FORBIDDEN", "egal", { status: 403 })), "HTTP 403 · FORBIDDEN");
  assert.equal(reasonOf(new ApiError("INTERNAL", "egal", { status: 500 })), "HTTP 500");
  assert.equal(reasonOf(new TypeError("Failed to fetch")), "TypeError: Failed to fetch");
  assert.equal(reasonOf(null), "ohne Angabe");
});

test("die Tickettypen kommen aus TANSS, nicht aus der Konfiguration", async () => {
  // Sie duerfen NICHT im Manifest stehen: Ein dort eingetragener Typ altert. Ein in TANSS
  // neu angelegter oder umbenannter kaeme nie an, ohne dass jemand ein Manifest erzeugt,
  // die Version erhoeht und neu ausrollt - und bis das bei allen Benutzern ankommt,
  // vergehen bis zu 72 Stunden.
  const client = {
    calls: [],
    get: async (pfad) => (pfad === "/api/v1/admin/ticketTypes"
      ? [{ id: 1, name: "Störung", active: true },
        { id: 2, name: "Anfrage", active: true },
        { id: 3, name: "Stillgelegt", active: false }]
      : []),
    put: async () => ({ content: {}, meta: {} }),
    post: async () => ({}),
    call: async () => [],
  };
  const options = await new TanssRepository({
    client,
    config: { ticketTypes: [{ id: 9, name: "Aus der Datei" }] },
  }).ticketOptions({});

  assert.deepEqual(options.types, [{ id: 1, name: "Störung" }, { id: 2, name: "Anfrage" }],
    "die API sticht die Konfiguration, und ein stillgelegter Typ wird nicht angeboten");
});

test("faellt die Typenroute aus, springt die Konfiguration ein und der Grund steht da", async () => {
  // Der Rueckfall ist fuer eine Installation gedacht, deren TANSS die Adminroute nicht
  // fuehrt - genau so, wie `defaults.supportTypeId` fuer die Leistungsarten einspringt.
  const client = {
    calls: [],
    get: async (pfad) => {
      if (pfad === "/api/v1/admin/ticketTypes") throw new ApiError("INTERNAL", "weg", { status: 404 });
      return [];
    },
    put: async () => ({ content: {}, meta: {} }),
    post: async () => ({}),
    call: async () => [],
  };
  const options = await new TanssRepository({
    client,
    config: { ticketTypes: [{ id: 9, name: "Aus der Datei" }] },
  }).ticketOptions({});

  assert.deepEqual(options.types, [{ id: 9, name: "Aus der Datei" }]);
  assert.ok(options.failures.some((f) => f.path === "/api/v1/admin/ticketTypes" && /404/.test(f.reason)));
});

test("ohne Route und ohne Konfiguration bleibt die Typenliste leer", async () => {
  // Dann verschwindet das Feld in der Maske. Ein Auswahlfeld ohne Auswahl sieht nach
  // einem Ausfall aus, und TANSS nimmt ein Ticket ohne Typ an.
  const client = {
    calls: [],
    get: async () => [],
    put: async () => ({ content: {}, meta: {} }),
    post: async () => ({}),
    call: async () => [],
  };
  const options = await new TanssRepository({ client, config: {} }).ticketOptions({});
  assert.deepEqual(options.types, []);
});

/* --------------------------------------------------- Abteilung am Techniker */

test("die Abteilungen eines Technikers werden als Menge gelesen", async () => {
  // In TANSS kann ein Mitarbeiter mehreren Abteilungen angehoeren: Neben dem Feld am
  // Mitarbeiter gibt es eine eigene Zuordnung. Wer nur das eine Feld naehme, boete einem
  // Techniker mit zwei Abteilungen nur eine an - und die zweite waere unerreichbar.
  const client = fakeClient({
    "GET /api/v1/employees/5": {
      content: { id: 5, departmentId: 1, departmentAssignments: [{ departmentId: 2 }, { departmentId: 1 }] },
    },
  });
  const ids = await new TanssRepository({ client, config: {} }).employeeDepartments(5);
  assert.deepEqual(ids.sort(), [1, 2]);
});

test("faellt der Abruf aus, bleibt die Menge leer und der Grund steht da", async () => {
  // Ausdruecklich KEIN Rueckfall auf die vollstaendige Liste: Das waere genau der
  // Zustand, der abgestellt werden soll - ein Ticket bei einer Abteilung, der der
  // Zugewiesene nicht angehoert -, nur ohne Hinweis darauf.
  const client = {
    calls: [],
    get: async () => { throw new ApiError("FORBIDDEN", "nein", { status: 403 }); },
    put: async () => ({}), post: async () => ({}), call: async () => [],
  };
  const failures = [];
  const ids = await new TanssRepository({ client, config: {} })
    .employeeDepartments(5, { failures });

  assert.deepEqual(ids, []);
  assert.ok(failures.some((f) => f.path === "/api/v1/employees/5" && /403/.test(f.reason)));
});

test("ohne Kennung wird gar nicht erst gefragt", async () => {
  const client = { calls: [], get: async () => { throw new Error("darf nicht"); },
    put: async () => ({}), post: async () => ({}), call: async () => [] };
  assert.deepEqual(await new TanssRepository({ client, config: {} }).employeeDepartments(null), []);
});

test("nennt die Antwort keine Abteilung, ist das Ergebnis leer - nicht alles", () => {
  assert.deepEqual(departmentIdsOf({ id: 5 }), []);
  assert.deepEqual(departmentIdsOf({ id: 5, departmentId: 0 }), []);
  assert.deepEqual(departmentIdsOf({ departments: [{ id: 7 }, 9] }).sort(), [7, 9]);
});
