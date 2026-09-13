/**
 * Die Sitzung ist die einzige Stelle, an der ein Token entsteht. Vier Zusagen haengen
 * an ihr, und jede einzelne faellt im Betrieb nur dann auf, wenn sie bricht:
 *
 * 1. Die Zugangsdaten stehen im Koerper, nie in der Adresse. In der Adresse landeten sie
 *    in jedem Proxy-Protokoll und im Verlauf des Browsers.
 * 2. `expire` ist ein Zeitstempel in SEKUNDEN. Als Millisekunden gelesen, gilt jedes
 *    Token als sofort abgelaufen, und das Pane erneuerte vor jedem Aufruf.
 * 3. Bei der Erneuerung geht NUR der Kopf `refreshToken` mit. Haengt ein Zugriffstoken
 *    daneben, antwortet TANSS gewoehnlich und erneuert nichts - die Sitzung liefe still
 *    ab, obwohl bei jedem Aufruf eine Erneuerung versucht wurde.
 * 4. Zwei gleichzeitige Erneuerungen teilen sich EINEN Aufruf. Sonst verbrauchen beide
 *    dasselbe Erneuerungstoken, die zweite scheitert, und der Techniker landet mitten in
 *    der Arbeit auf der Anmeldeseite.
 */

import test from "node:test";
import assert from "node:assert/strict";

import { Session } from "../../taskpane/js/tanss/session.js";

const BASE = "https://tanss.example/backend";

/** Ein Speicher, wie ihn der Browser stellt - mehr braucht die Sitzung nicht. */
function fakeStorage() {
  const data = new Map();
  return {
    getItem: (key) => (data.has(key) ? data.get(key) : null),
    setItem: (key, value) => data.set(key, String(value)),
    removeItem: (key) => data.delete(key),
  };
}

/** Eine frische Sitzung mit aufgezeichneten Aufrufen. */
function sessionWith(replies) {
  globalThis.localStorage = fakeStorage();
  const calls = [];
  const queue = [...replies];
  const session = new Session({
    baseUrl: BASE,
    fetchImpl: async (url, init) => {
      calls.push({ url: new URL(url), init });
      const reply = queue.length > 1 ? queue.shift() : queue[0];
      return new Response(reply.body === undefined ? "" : reply.body, {
        status: reply.status || 200,
        headers: { "Content-Type": "application/json" },
      });
    },
  });
  session.clear();
  return { session, calls };
}

const loginReply = (overrides = {}) =>
  JSON.stringify({
    content: {
      employeeId: 42,
      apiKey: "Bearer neu",
      expire: Math.floor(Date.now() / 1000) + 4 * 3600,
      refresh: "Bearer erneuerung",
      employeeType: "TECHNICAN",
      ...overrides,
    },
  });

test("die Zugangsdaten stehen im Koerper, nie in der Adresse", async () => {
  const { session, calls } = sessionWith([{ body: loginReply() }]);
  await session.login({ username: "tester", password: "geheim" });

  const { url, init } = calls[0];
  assert.equal(url.search, "", "die Adresse darf keine Parameter tragen");
  assert.equal(init.method, "POST");
  assert.deepEqual(JSON.parse(init.body), { username: "tester", password: "geheim" });
});

test("die Einmalkennung wandert mit, wenn es eine gibt", async () => {
  const { session, calls } = sessionWith([{ body: loginReply() }]);
  await session.login({ username: "a", password: "b", token: "123456" });
  assert.equal(JSON.parse(calls[0].init.body).token, "123456");
});

test("ohne Einmalkennung wird das Feld nicht gesendet", async () => {
  const { session, calls } = sessionWith([{ body: loginReply() }]);
  await session.login({ username: "a", password: "b" });
  assert.equal("token" in JSON.parse(calls[0].init.body), false);
});

test("expire ist in Sekunden und macht das Token nicht sofort ungueltig", async () => {
  const { session } = sessionWith([{ body: loginReply() }]);
  await session.login({ username: "a", password: "b" });
  assert.equal(session.isFresh(), true);
  assert.equal(session.employeeId(), 42);
  assert.equal(session.apiToken(), "Bearer neu");
});

test("eine abgelehnte Anmeldung nennt alle drei moeglichen Ursachen", async () => {
  const { session } = sessionWith([{ status: 403, body: JSON.stringify({ content: {} }) }]);
  await assert.rejects(
    () => session.login({ username: "a", password: "falsch" }),
    (error) => {
      assert.equal(error.code, "UNAUTHENTICATED");
      assert.match(error.message, /Benutzername, Kennwort oder Einmalkennung/);
      return true;
    },
  );
  assert.equal(session.exists(), false);
});

test("eine Anmeldung ohne Mitarbeiternummer gilt nicht als Anmeldung", async () => {
  const { session } = sessionWith([{ body: loginReply({ employeeId: 0 }) }]);
  await assert.rejects(
    () => session.login({ username: "a", password: "b" }),
    (error) => error.code === "UNAUTHENTICATED",
  );
});

test("bei der Erneuerung geht NUR der refreshToken-Kopf mit", async () => {
  const { session, calls } = sessionWith([
    { body: loginReply({ expire: Math.floor(Date.now() / 1000) - 10 }) },
    { body: loginReply({ apiKey: "Bearer frisch" }) },
  ]);
  await session.login({ username: "a", password: "b" });
  assert.equal(session.isFresh(), false, "das Token muss als abgelaufen gelten");

  await session.ensureFresh();

  const renew = calls[1];
  const headers = new Headers(renew.init.headers);
  assert.equal(headers.get("refreshToken"), "Bearer erneuerung");
  assert.equal(headers.has("apiToken"), false, "ein Zugriffstoken verhinderte die Erneuerung");
  assert.equal(session.apiToken(), "Bearer frisch");
});

test("zwei gleichzeitige Erneuerungen ergeben genau einen Aufruf", async () => {
  const { session, calls } = sessionWith([
    { body: loginReply({ expire: Math.floor(Date.now() / 1000) - 10 }) },
    { body: loginReply({ apiKey: "Bearer frisch" }) },
  ]);
  await session.login({ username: "a", password: "b" });

  await Promise.all([session.ensureFresh(), session.ensureFresh(), session.ensureFresh()]);
  assert.equal(calls.length, 2, "Anmeldung plus genau eine Erneuerung");
});

test("antwortet die Route gewoehnlich, gilt die Erneuerung als gescheitert", async () => {
  // Kein `apiKey` in der Antwort: Der Kopf wurde nicht ausgewertet. Das Pane darf dann
  // nicht weiterlaufen, als sei erneuert worden.
  const { session } = sessionWith([
    { body: loginReply({ expire: Math.floor(Date.now() / 1000) - 10 }) },
    { body: JSON.stringify({ content: { ownCompanyId: 1, rights: [] } }) },
  ]);
  await session.login({ username: "a", password: "b" });

  await assert.rejects(
    () => session.ensureFresh(),
    (error) => error.code === "UNAUTHENTICATED",
  );
  assert.equal(session.exists(), false, "die unbrauchbare Sitzung wird verworfen");
});

test("ein frisches Token wird nicht ohne Not erneuert", async () => {
  const { session, calls } = sessionWith([{ body: loginReply() }]);
  await session.login({ username: "a", password: "b" });
  await session.ensureFresh();
  assert.equal(calls.length, 1, "nur die Anmeldung");
});

test("ohne Erneuerungstoken gibt es nichts zu erneuern", async () => {
  const { session } = sessionWith([{ body: loginReply({ refresh: "" }) }]);
  await session.login({ username: "a", password: "b" });
  await assert.rejects(() => session.renew(), (error) => error.code === "UNAUTHENTICATED");
});

test("ein gesperrter Speicher nimmt der Sitzung nicht die Funktion", async () => {
  const { session } = sessionWith([{ body: loginReply() }]);
  globalThis.localStorage = {
    getItem: () => {
      throw new Error("gesperrt");
    },
    setItem: () => {
      throw new Error("gesperrt");
    },
    removeItem: () => {},
  };
  await session.login({ username: "a", password: "b" });
  assert.equal(session.employeeId(), 42, "der Arbeitsspeicher traegt die Sitzung");
});
