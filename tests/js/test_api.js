/**
 * Der Weg von der Seite bis zur TANSS-Anfrage, in einem Stueck.
 *
 * Die Seiten rufen `api.get("api/me")` und lesen `result.data`. Zwischen beidem liegen
 * jetzt Verteiler, Laufzeit, Sitzung, Fachschicht und Client. Diese Datei prueft den
 * ganzen Zug, weil die Fehler, auf die es ankommt, genau an den Naehten entstehen:
 *
 * - Es darf NICHTS fliegen. Eine Seite, die einen Wurf bekommt statt eines Umschlags,
 *   bleibt fuer immer im Ladezustand.
 * - Jede /api/v1-Anfrage traegt den Handelnden. Ohne ihn antwortet TANSS mit 403, und
 *   eine leere 403 sieht in der Oberflaeche aus wie "keine Treffer".
 * - Ein Doppelklick darf kein zweites Ticket anlegen.
 * - Sobald das Ticket existiert, darf kein Fehler den Vorgang mehr scheitern lassen -
 *   sonst kennt der Techniker die Nummer nicht und legt beim zweiten Versuch ein
 *   zweites an.
 *
 * Die Reihenfolge der Faelle ist Absicht: Die Laufzeit ist ein Modulzustand, und die
 * Anmeldung in der Mitte gilt fuer alles danach - genau wie im Pane.
 */

import test from "node:test";
import assert from "node:assert/strict";

import { TanssRepository } from "../../taskpane/js/tanss/repository.js";

import * as api from "../../taskpane/js/api.js";

const API_BASE = "https://tanss.example/backend";

/** Was die Attrappe geantwortet hat, und worauf. */
const calls = [];

/** Antworten je Pfad. Ein Eintrag darf eine Funktion sein, um zu zaehlen. */
let routes = {};

function reply(body, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

function fakeStorage() {
  const data = new Map();
  return {
    getItem: (key) => (data.has(key) ? data.get(key) : null),
    setItem: (key, value) => data.set(key, String(value)),
    removeItem: (key) => data.delete(key),
  };
}

globalThis.localStorage = fakeStorage();
globalThis.location = { href: "https://addon.example/", host: "addon.example", hostname: "addon.example" };
globalThis.document = { baseURI: "https://addon.example/" };

globalThis.fetch = async (input, init = {}) => {
  const url = new URL(String(input), "https://addon.example/");
  const path = url.pathname.replace("/backend", "");
  calls.push({ url, path, method: (init.method || "GET").toUpperCase(), init });

  if (url.pathname === "/config.json") {
    return reply({
      apiBase: API_BASE,
      frontendBase: "https://tanss.example",
      instance: "pruefstand",
      ticketTypes: [{ id: 1, name: "Stoerung" }],
    });
  }
  const handler = routes[`${(init.method || "GET").toUpperCase()} ${path}`];
  if (!handler) return reply({ error: { text: "OBJECT_NOT_FOUND" } }, 404);
  return typeof handler === "function" ? handler() : reply(handler);
};

const login = () => reply({
  content: {
    employeeId: 42,
    apiKey: "Bearer tok",
    expire: Math.floor(Date.now() / 1000) + 4 * 3600,
    refresh: "Bearer erneuerung",
  },
});

/* ------------------------------------------------------- Vor der Anmeldung */

test("ein Pfad, den es nicht gibt, endet als Umschlag und nicht als Wurf", async () => {
  const result = await api.get("api/gibtesnicht");
  assert.equal(result.ok, false);
  assert.equal(result.error.code, "INTERNAL");
  assert.equal(typeof result.error.message, "string");
  assert.ok(result.error.message.length > 0, "ein Fehlerband ohne Text hilft niemandem");
});

test("ohne Anmeldung wird nicht gehandelt, und es fliegt trotzdem nichts", async () => {
  routes = {};
  const result = await api.get("api/me");
  assert.equal(result.ok, false);
  assert.equal(result.error.code, "UNAUTHENTICATED");
  // Keine einzige TANSS-Anfrage darf abgesetzt worden sein.
  assert.equal(calls.some((c) => c.path.startsWith("/api/v1/")), false);
});

/* --------------------------------------------------------------- Anmeldung */

test("die Anmeldung braucht selbst kein Token und traegt keinen Handelnden", async () => {
  routes = { "POST /api/v1/login": login };
  calls.length = 0;

  const result = await api.post("api/auth/tanss", {
    json: { username: "tester", password: "geheim" },
  });

  assert.equal(result.ok, true);
  assert.equal(result.data.employeeId, 42);

  const anmeldung = calls.find((c) => c.path === "/api/v1/login");
  assert.equal(anmeldung.url.searchParams.has("loggedInUserId"), false);
  assert.deepEqual(JSON.parse(anmeldung.init.body), { username: "tester", password: "geheim" });
});

/* ------------------------------------------------------- Nach der Anmeldung */

test("jede fachliche Anfrage traegt den Handelnden aus der Sitzung", async () => {
  routes = {
    "PUT /api/v1/search": { content: { companies: [{ id: 3, name: "Muster GmbH" }] } },
  };
  calls.length = 0;

  const result = await api.get("api/search/companies", { query: { q: "muster" } });
  assert.equal(result.ok, true);
  assert.equal(result.data.items[0].name, "Muster GmbH");

  const suche = calls.find((c) => c.path === "/api/v1/search");
  assert.equal(suche.url.searchParams.get("loggedInUserId"), "42");
});

test("eine zu kurze Suche kostet keinen Aufruf", async () => {
  calls.length = 0;
  const result = await api.get("api/search/companies", { query: { q: "a" } });
  assert.equal(result.ok, false);
  assert.equal(result.error.code, "CONFLICT");
  assert.match(result.error.message, /mindestens 2 Zeichen/);
  assert.equal(calls.length, 0);
});

test("die Identitaet kommt aus der Anmeldung und nicht aus einer Zuordnung", async () => {
  routes = {
    "GET /api/v1/employees/technicians": [
      { id: 42, name: "Anna Muster", listType: "TECHNICIAN" },
    ],
  };
  const result = await api.get("api/me");
  assert.equal(result.ok, true);
  assert.equal(result.data.employeeId, 42);
  assert.equal(result.data.name, "Anna Muster");
  assert.equal(result.data.instance, "pruefstand");
});

test("die Technikerliste wird zwischengespeichert", async () => {
  // Sie aendert sich im Tagesbetrieb nicht, und der Abruf kostet bei jedem Oeffnen des
  // Panes eine Umlaufzeit. Der Preis steht in der Fachschicht: Ein neuer Kollege ist
  // hoechstens eine Viertelstunde lang unsichtbar - und ein AUSFALL faellt ebenso lange
  // nicht auf. Dieser Fall haelt beide Seiten der Abwaegung fest.
  routes = {};
  calls.length = 0;

  const result = await api.get("api/me");
  assert.equal(result.ok, true);
  assert.equal(result.data.name, "Anna Muster", "aus dem Zwischenspeicher");
  assert.equal(calls.some((c) => c.path === "/api/v1/employees/technicians"), false,
    "es wurde gar nicht erst gefragt");
  assert.deepEqual(result.data.warnings, []);
});

/* ------------------------------------------------------------ Ticketanlage */

/** Ein Formular, wie die Ticketmaske es absendet. */
function ticketForm(overrides = {}, mail = null) {
  const form = new FormData();
  form.set("ticket", JSON.stringify({
    companyId: 3, title: "Drucker streikt", content: "Text", ...overrides,
  }));
  if (mail) {
    form.set("eml", new Blob(["nachricht"]), "nachricht.eml");
    form.set("internetMessageId", mail.messageId || "");
    form.set("subject", mail.subject || "");
  }
  return form;
}

const OPTIONEN = {
  "GET /api/v1/tickets/": { content: {}, meta: {} },
  "GET /api/v1/admin/ticketStates": { content: [] },
  "GET /api/v1/employees/technicians": { content: [] },
  "GET /api/v1/employees/departments": { content: [] },
};

test("ohne Titel entsteht kein Ticket", async () => {
  routes = { ...OPTIONEN };
  const result = await api.post("api/tickets", { form: ticketForm({ title: "  " }) });
  assert.equal(result.ok, false);
  assert.equal(result.error.code, "TITLE_REQUIRED");
});

test("ohne Melderzwang laesst sich ein Ticket ohne Melder anlegen", () => {
  // Frueher stand hier der umgekehrte Fall: Die Feldsteuerung meldete einen Melderzwang,
  // und der Vorgang wurde vor dem Schreiben abgewiesen. Die Route, aus der dieser Zwang
  // gelesen wurde, gibt es nicht - unter /api/v1/tickets sind nur POST und PUT definiert.
  // Der Zwang konnte also nie greifen; die Pruefung in `validate` bleibt trotzdem stehen,
  // weil sie zutraefe, sobald eine Quelle dafuer existiert.
  //
  // Was hier festgehalten wird, ist die geltende Annahme: KEIN Zwang. Sie ist die
  // mildere - ein erfundener Zwang hielte den Techniker von einem gueltigen Ticket ab,
  // waehrend eine fehlende Pflicht TANSS beim Anlegen selbst auffaellt.
  assert.equal(TanssRepository.FIELD_RULES.remitterRequired, false);
  assert.equal(TanssRepository.FIELD_RULES.forceAssignment, false);
  assert.equal("autoAssignedEmployeeId" in TanssRepository.FIELD_RULES, false,
    "das Feld kommt in TANSS nicht vor und darf keine Vorbelegung erzeugen");
});
test("ein Doppelklick legt kein zweites Ticket an", async () => {
  let angelegt = 0;
  routes = {
    ...OPTIONEN,
    "POST /api/v1/tickets": () => {
      angelegt += 1;
      return reply({ content: { id: 4711, companyId: 3 }, meta: {} });
    },
  };

  const key = api.newIdempotencyKey();
  const form = ticketForm();
  const [erste, zweite] = await Promise.all([
    api.post("api/tickets", { form, idempotencyKey: key }),
    api.post("api/tickets", { form, idempotencyKey: key }),
  ]);

  assert.equal(erste.ok, true);
  assert.equal(erste.data.ticketId, 4711);
  assert.deepEqual(zweite, erste, "der zweite Klick bekommt dieselbe Antwort");
  assert.equal(angelegt, 1, "und es wurde genau einmal angelegt");
});

test("scheitert die Mailablage, bleibt die Ticketnummer trotzdem erhalten", async () => {
  routes = {
    ...OPTIONEN,
    "POST /api/v1/tickets": { content: { id: 4712, companyId: 3 }, meta: {} },
    // Kein Upload-Pfad: Die Ablage schlaegt fehl.
  };

  const result = await api.post("api/tickets", {
    form: ticketForm({}, { messageId: "<a@b>", subject: "Drucker" }),
    idempotencyKey: api.newIdempotencyKey(),
  });

  assert.equal(result.ok, true, "der Vorgang darf nicht als Fehler enden");
  assert.equal(result.data.ticketId, 4712);
  assert.equal(result.data.mail.status, "failed");
  assert.ok(result.data.mail.cause, "die Ursache entscheidet, ob ein zweiter Versuch hilft");
});

test("die Ticketadresse entsteht aus der eingestellten Oberflaeche", async () => {
  routes = {
    ...OPTIONEN,
    "POST /api/v1/tickets": { content: { id: 4713, companyId: 3 }, meta: {} },
  };
  const result = await api.post("api/tickets", {
    form: ticketForm(), idempotencyKey: api.newIdempotencyKey(),
  });
  assert.equal(result.data.url, "https://tanss.example/index.php?section=bug&sub=view&bugID=4713");
});

/* -------------------------------------------------------------- Mailablage */

test("eine bereits angehaengte Nachricht wird gewarnt, nicht gesperrt", async () => {
  let hochgeladen = 0;
  routes = {
    "GET /api/v1/tickets/history/500": { content: { mails: [] } },
    "GET /api/v1/tickets/500/documents": { content: [] },
    "POST /api/v1/tickets/500/upload": () => {
      hochgeladen += 1;
      return reply({ content: [{ status: "OK", name: "a.eml", size: 9 }] });
    },
  };

  const form = () => {
    const f = new FormData();
    f.set("eml", new Blob(["x"]), "a.eml");
    f.set("internetMessageId", "<doppelt@example.de>");
    return f;
  };

  const erste = await api.post("api/tickets/500/mail", { form: form() });
  assert.equal(erste.ok, true);
  assert.equal(hochgeladen, 1);

  const zweite = await api.post("api/tickets/500/mail", { form: form() });
  assert.equal(zweite.ok, false);
  assert.equal(zweite.error.code, "MAIL_ALREADY_ATTACHED");
  assert.equal(hochgeladen, 1, "die Warnung hat den zweiten Upload verhindert");

  // Mit ausdruecklichem Willen geht es trotzdem - es gibt berechtigte Gruende.
  const erzwungen = form();
  erzwungen.set("force", "true");
  const dritte = await api.post("api/tickets/500/mail", { form: erzwungen });
  assert.equal(dritte.ok, true);
  assert.equal(hochgeladen, 2);
});

test("eine leer angekommene Nachricht bekommt ihren eigenen Satz", async () => {
  // TANSS kennt vier Zustaende je hochgeladener Datei: OK, TOO_BIG, ERROR und EMPTY. Das
  // Pane benannte zwei; die anderen beiden fielen in "Im Dienst ist ein unerwarteter
  // Fehler aufgetreten". "EMPTY" nennt aber eine Ursache, die der Techniker sofort
  // versteht und selbst beheben kann - sie gehoert nicht hinter einen Sammelsatz.
  routes = {
    "GET /api/v1/tickets/history/500": { content: { mails: [] } },
    "GET /api/v1/tickets/500/documents": { content: [] },
    "POST /api/v1/tickets/500/upload": { content: [{ status: "EMPTY", name: "a.eml", size: 0 }] },
  };

  const f = new FormData();
  f.set("eml", new Blob([""]), "a.eml");
  f.set("internetMessageId", "<leer@example.de>");

  const result = await api.post("api/tickets/500/mail", { form: f });
  assert.equal(result.ok, false);
  assert.equal(result.error.code, "MAIL_EMPTY");
  assert.match(result.error.message, /ohne Inhalt/);
});

test("ein unbekannter Zustand faellt weiterhin auf den Sammelfall", async () => {
  // Absicht: Ein Zustand, den wir nicht kennen, soll als unerwarteter Fehler auffallen -
  // und nicht stillschweigend als Erfolg durchgehen.
  routes = {
    "GET /api/v1/tickets/history/500": { content: { mails: [] } },
    "GET /api/v1/tickets/500/documents": { content: [] },
    "POST /api/v1/tickets/500/upload": { content: [{ status: "WAS_AUCH_IMMER" }] },
  };

  const f = new FormData();
  f.set("eml", new Blob(["x"]), "a.eml");
  f.set("internetMessageId", "<neu@example.de>");

  const result = await api.post("api/tickets/500/mail", { form: f });
  assert.equal(result.ok, false);
  assert.equal(result.error.code, "INTERNAL");
});
