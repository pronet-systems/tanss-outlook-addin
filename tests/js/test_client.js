/**
 * Der TANSS-Client haelt zwei Regeln, die keine Verabredung sein duerfen, weil ihr
 * Bruch nicht auffaellt:
 *
 * 1. `/api/v1/**` traegt IMMER `loggedInUserId`. Fehlt er, antwortet TANSS mit 403 -
 *    und eine leere 403 sieht in der Oberflaeche aus wie "keine Treffer".
 * 2. `/api/tanss.x/v1/**` traegt ihn NIEMALS. Mit dem Parameter nimmt der Server
 *    andere Zweige als die erprobten; der Unterschied faellt erst im Betrieb auf.
 *
 * Dazu die Auswertung: Status vor Koerper, und eine leere Antwort ist nur bei Erfolg
 * eine leere Antwort.
 */

import test from "node:test";
import assert from "node:assert/strict";

import { TanssClient, codeFor } from "../../taskpane/js/tanss/client.js";
import { ApiError } from "../../taskpane/js/tanss/errors.js";

const BASE = "https://tanss.example/backend";

/** Ein Client, der jede Anfrage aufzeichnet und eine vorgegebene Antwort liefert. */
function clientWith(replies, { employeeId = 42 } = {}) {
  const calls = [];
  const queue = Array.isArray(replies) ? [...replies] : [replies];
  globalThis.fetch = async (url, init) => {
    calls.push({ url: new URL(url), init });
    const reply = queue.length > 1 ? queue.shift() : queue[0];
    const status = reply.status || 200;
    // 204, 205 und 304 duerfen keinen Koerper tragen - auch keinen leeren. Ein
    // Leerstring an dieser Stelle laesst `Response` werfen, und der Client saehe einen
    // Netzfehler statt der Antwort, die geprueft werden soll.
    const bodyless = status === 204 || status === 205 || status === 304;
    return new Response(bodyless ? null : (reply.body === undefined ? "" : reply.body), {
      status,
      headers: reply.headers || { "Content-Type": "application/json" },
    });
  };
  const client = new TanssClient({
    baseUrl: BASE,
    tokens: { apiToken: () => "Bearer tok" },
    employeeId: () => employeeId,
  });
  return { client, calls };
}

const envelope = (content, meta) =>
  JSON.stringify(meta === undefined ? { content } : { content, meta });

test("eine /api/v1-Route traegt loggedInUserId", async () => {
  const { client, calls } = clientWith({ body: envelope([]) });
  await client.get("/api/v1/employees/ownState");
  assert.equal(calls[0].url.searchParams.get("loggedInUserId"), "42");
});

test("eine /api/tanss.x/v1-Route traegt loggedInUserId NIEMALS", async () => {
  const { client, calls } = clientWith({ body: envelope([]) });
  await client.get("/api/tanss.x/v1/technicians");
  assert.equal(calls[0].url.searchParams.has("loggedInUserId"), false);
});

test("die Anmeldung ist die einzige /api/v1-Route ohne Handelnden", async () => {
  const { client, calls } = clientWith({ body: envelope({ employeeId: 1 }) });
  await client.call("POST", "/api/v1/login", { json: { username: "a" } });
  assert.equal(calls[0].url.searchParams.has("loggedInUserId"), false);
});

test("ohne angemeldeten Mitarbeiter wird gar nicht erst gesendet", async () => {
  const { client, calls } = clientWith({ body: envelope([]) }, { employeeId: 0 });
  await assert.rejects(
    () => client.get("/api/v1/tickets"),
    (error) => error instanceof ApiError && error.code === "UNAUTHENTICATED",
  );
  assert.equal(calls.length, 0, "es darf keine Anfrage abgesetzt worden sein");
});

test("ein mitgegebenes loggedInUserId ist ein Programmfehler, kein Datenfehler", async () => {
  const { client } = clientWith({ body: envelope([]) });
  await assert.rejects(
    () => client.get("/api/v1/tickets", { query: { loggedInUserId: 7 } }),
    (error) => error instanceof ApiError && error.code === "INTERNAL",
  );
});

test("das Token geht in apiToken, nicht in Authorization", async () => {
  const { client, calls } = clientWith({ body: envelope([]) });
  await client.get("/api/v1/employees/ownState");
  const headers = new Headers(calls[0].init.headers);
  assert.equal(headers.get("apiToken"), "Bearer tok");
  assert.equal(headers.has("Authorization"), false);
});

test("der Umschlag wird ausgepackt, meta auf Wunsch mitgeliefert", async () => {
  const { client } = clientWith({ body: envelope([{ id: 1 }], { properties: {} }) });
  const plain = await client.get("/api/v1/tickets");
  assert.deepEqual(plain, [{ id: 1 }]);

  const { client: second } = clientWith({ body: envelope([{ id: 1 }], { a: 1 }) });
  const pair = await second.get("/api/v1/tickets", { wantMeta: true });
  assert.deepEqual(pair, { content: [{ id: 1 }], meta: { a: 1 } });
});

test("eine leere 403 ist ein Fehler und niemals eine leere Trefferliste", async () => {
  const { client } = clientWith({ status: 403, body: "" });
  await assert.rejects(
    () => client.get("/api/v1/tickets"),
    (error) => error instanceof ApiError && error.code === "UNAUTHENTICATED",
  );
});

test("eine leere 204 ist ein leerer Erfolg", async () => {
  const { client } = clientWith({ status: 204, body: "" });
  assert.deepEqual(await client.get("/api/v1/tickets"), {});
});

test("ein HTTP 200 mit einer Anmeldeseite darin gilt nicht als Erfolg", async () => {
  const { client } = clientWith({
    status: 200,
    body: "<html>Bitte anmelden</html>",
    headers: { "Content-Type": "text/html" },
  });
  await assert.rejects(
    () => client.get("/api/v1/tickets"),
    (error) => error.code === "CLIENT_BAD_RESPONSE",
  );
});

test("ein Lizenzproblem bleibt ein Lizenzproblem, auch als HTTP 500", () => {
  const { code } = codeFor(500, { error: { licenseError: "NO_BASE_LICENSE_FOUND" } });
  assert.equal(code, "TANSS_LICENSE_ERROR");
});

test("der Fehlername schlaegt den Status", () => {
  assert.equal(codeFor(400, { error: { text: "INVALID_EML" } }).code, "MAIL_UNPARSABLE");
  assert.equal(
    codeFor(500, { error: { text: "CHANGES_WERE_DISCARDED" } }).code,
    "APPOINTMENT_READONLY",
  );
  assert.equal(
    codeFor(403, { error: { text: "FORBIDDEN_NO_COMPANY_ACCESS" } }).code,
    "FORBIDDEN_NO_COMPANY_ACCESS",
  );
});

test("ein unerreichbares TANSS wird nicht als Programmfehler ausgegeben", () => {
  assert.equal(codeFor(502, null).code, "TANSS_UNAVAILABLE");
  assert.equal(codeFor(503, null).code, "TANSS_UNAVAILABLE");
});
