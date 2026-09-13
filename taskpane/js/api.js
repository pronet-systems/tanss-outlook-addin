/**
 * Der Verteiler zwischen Oberflaeche und Fachschicht.
 *
 * Frueher war das eine Huelle um `fetch` gegen einen eigenen Dienst. Den Dienst gibt es
 * nicht mehr; die Vorgaenge laufen jetzt im Pane. Die SCHNITTSTELLE ist absichtlich
 * dieselbe geblieben - `get`, `post`, `put` mit denselben Pfaden und demselben
 * Umschlag. Dadurch sind die Seiten unveraendert geblieben, und der Umbau laesst sich
 * an einer Stelle nachlesen statt an zwoelf.
 *
 * Ein Umschlag, nicht zwei. Zwei nebeneinander gefuehrte Fehlerformen enden im selben
 * Fehler: Der Leser greift auf ein Feld zu, das diesmal fehlt, wirft im eigenen `catch`,
 * und die Seite bleibt fuer immer im Ladezustand. Hier gibt jede Funktion
 * `{ok:true, data}` oder `{ok:false, error:{code, message}}` zurueck - auch bei
 * Netzabbruch, Zeitueberschreitung und abgelaufener Anmeldung. Es fliegt nichts.
 *
 * Die Idempotenz liegt jetzt ebenfalls hier, und sie kann weniger als vorher: Sie
 * verhindert den Doppelklick INNERHALB dieses Panes. Ein zweiter Versuch aus einem
 * neuen Fenster ist davon nicht mehr gedeckt - dafuer braeuchte es einen gemeinsamen
 * Zustand, und den gibt es ohne Dienst nicht. Die Maske sperrt deshalb zusaetzlich
 * ihren eigenen Knopf, solange ein Vorgang laeuft.
 */

import { errorText } from "../i18n/de.js";
import { ApiError } from "./tanss/errors.js";
import { ready, start } from "./runtime.js";
import * as appointments from "./services/appointments.js";
import * as meta from "./services/meta.js";
import * as search from "./services/search.js";
import * as tickets from "./services/tickets.js";

/* ------------------------------------------------------------------ Umschlag */

function ok(data) {
  return { ok: true, data, status: 200, traceId: "" };
}

/**
 * Fehlerumschlag. `message` ist immer gefuellt - lieber ein allgemeiner Satz als ein
 * Fehlerband ohne Text, vor dem der Benutzer ratlos sitzt.
 */
function fail(code, message = "", extra = {}) {
  const error = { code, message: errorText({ code, message }) };
  if (extra.detail !== undefined) error.detail = extra.detail;
  return { ok: false, error, status: 0, traceId: "" };
}

/**
 * Idempotenzschluessel.
 *
 * `crypto.randomUUID` gibt es in jedem sicheren Kontext, und das Pane laeuft
 * ausschliesslich ueber https. Der Rueckfall deckt aeltere Webviews ab; er benutzt
 * weiterhin den Zufallsgenerator des Systems, nie `Math.random` - ein erratbarer
 * Schluessel liesse einen fremden Aufrufer die Antwort eines anderen Vorgangs abholen.
 */
export function newIdempotencyKey() {
  const crypto = globalThis.crypto;
  if (crypto && typeof crypto.randomUUID === "function") return crypto.randomUUID();
  const bytes = new Uint8Array(16);
  crypto.getRandomValues(bytes);
  bytes[6] = (bytes[6] & 0x0f) | 0x40;
  bytes[8] = (bytes[8] & 0x3f) | 0x80;
  const hex = Array.from(bytes, (b) => b.toString(16).padStart(2, "0")).join("");
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
}

/**
 * Laufende und abgeschlossene Vorgaenge je Idempotenzschluessel.
 *
 * Ein zweiter Aufruf mit demselben Schluessel bekommt DASSELBE Promise - und damit
 * dieselbe Antwort, nicht einen zweiten Vorgang. Genau das verhindert, dass ein
 * Doppelklick zwei Tickets anlegt.
 */
const inflight = new Map();

/* ------------------------------------------------------------------ Routen */

/**
 * Die Vorgaenge, auf die die Seiten zugreifen.
 *
 * `auth` sagt, ob vor dem Aufruf ein gueltiges Token vorliegen muss. Die Anmeldung
 * selbst ist der einzige Vorgang, der ohne auskommt - sie stellt das Token ja erst her.
 */
const ROUTES = [
  ["GET", /^api\/me$/, meta.me, { auth: true }],
  ["POST", /^api\/prefs$/, (m, body) => meta.setPref(body), { auth: false }],
  ["POST", /^api\/client-error$/, (m, body) => meta.reportClientError(body), { auth: false }],
  ["POST", /^api\/auth\/tanss$/, (m, body) => meta.login(body), { auth: false }],
  ["POST", /^api\/auth\/logout$/, () => meta.logout(), { auth: false }],

  ["POST", /^api\/mail\/context$/, (m, body, o) => tickets.mailContext(body, o), { auth: true }],
  ["GET", /^api\/tickets\/options$/, (m, q, o) => tickets.ticketOptions(q, o), { auth: true }],
  ["POST", /^api\/tickets$/, (m, form, o) => tickets.createTicket(form, o), { auth: true }],
  ["POST", /^api\/tickets\/(\d+)\/mail$/, (m, form, o) => tickets.attachMail(Number(m[1]), form, o), { auth: true }],

  ["GET", /^api\/search\/companies$/, (m, q, o) => search.searchCompanies(q, o), { auth: true }],
  ["GET", /^api\/search\/employees$/, (m, q, o) => search.searchEmployees(q, o), { auth: true }],
  ["GET", /^api\/search\/tickets$/, (m, q, o) => search.searchTickets(q, o), { auth: true }],
  ["POST", /^api\/companies\/(\d+)\/contacts$/, (m, body, o) => search.createContact(Number(m[1]), body, o), { auth: true }],

  ["GET", /^api\/appointments\/resolve$/, (m, q, o) => appointments.resolve(q, o), { auth: true }],
  ["GET", /^api\/appointments\/options$/, (m, q, o) => appointments.options(q, o), { auth: true }],
  ["POST", /^api\/appointments$/, (m, body, o) => appointments.create(body, o), { auth: true }],
  ["PUT", /^api\/appointments\/(\d+)$/, (m, body, o) => appointments.save(Number(m[1]), body, o), { auth: true }],
];

/** Findet den Vorgang zu Methode und Pfad. */
function route(method, path) {
  const clean = String(path).replace(/^\/+/, "").split("?")[0];
  for (const [verb, pattern, handler, flags] of ROUTES) {
    if (verb !== method) continue;
    const match = pattern.exec(clean);
    if (match) return { handler, flags, match };
  }
  return null;
}

/* ---------------------------------------------------------------- Hauptweg */

/**
 * Ein Vorgang.
 *
 * @param {string} method
 * @param {string} path Wie bisher, z. B. `api/tickets`.
 * @param {{query?: object, json?: object, form?: FormData, signal?: AbortSignal,
 *          idempotencyKey?: string}} [options]
 * @returns {Promise<{ok: true, data: any}|{ok: false, error: {code: string, message: string}}>}
 */
export async function request(method, path, options = {}) {
  const upper = String(method).toUpperCase();
  const found = route(upper, path);
  if (!found) {
    // Ein Pfad, den es nicht gibt, ist ein Programmfehler - aber auch er darf die
    // Seite nicht im Ladezustand stehen lassen.
    return fail("INTERNAL", "", { detail: `${upper} ${path}` });
  }

  const key = options.idempotencyKey;
  if (key && inflight.has(key)) {
    // Derselbe Vorgang, zweiter Klick: dieselbe Antwort, kein zweites Ticket.
    return inflight.get(key);
  }

  const promise = run(found, options);
  if (key) {
    inflight.set(key, promise);
    // Der Eintrag bleibt: Ein spaeterer Klick mit demselben Schluessel soll die alte
    // Antwort wiederbekommen und nicht einen neuen Vorgang starten. Er lebt nur so
    // lange wie das Pane - laenger braucht es nicht, denn mit dem Pane verschwindet
    // auch das Formular, aus dem der Schluessel stammt.
  }
  return promise;
}

async function run({ handler, flags, match }, options) {
  try {
    if (flags.auth) await ready();
    else await start();

    const payload = options.form !== undefined
      ? options.form
      : (options.json !== undefined ? options.json : (options.query || {}));

    const data = await handler(match, payload, { signal: options.signal });
    return ok(data === undefined ? null : data);
  } catch (error) {
    if (options.signal && options.signal.aborted) return fail("CLIENT_UNEXPECTED");
    if (error instanceof ApiError) {
      return fail(error.code, error.message === error.code ? "" : error.message,
        { detail: error.detail || undefined });
    }
    if (error && error.name === "TypeError") return fail("CLIENT_NETWORK");
    return fail("INTERNAL", "", { detail: error && error.name });
  }
}

export function get(path, options = {}) {
  return request("GET", path, options);
}

export function post(path, options = {}) {
  return request("POST", path, options);
}

export function put(path, options = {}) {
  return request("PUT", path, options);
}

/**
 * Meldet einen unerwarteten Clientfehler.
 *
 * Er geht an KEINEN Empfaenger ausser der Konsole dieses Rechners. Ein Fehlerbericht an
 * einen Dritten waere ein weiterer Empfaenger fuer Daten, die durch dieses Pane
 * laufen - Betreff, Klartext, Token.
 */
let reporting = false;

export async function reportClientError(route_, message) {
  if (reporting) return;
  reporting = true;
  try {
    await post("api/client-error", {
      json: {
        route: String(route_ || "").slice(0, 100),
        message: String(message || "").slice(0, 500),
      },
    });
  } catch {
    // Absichtlich stumm: Ein Fehlerbericht, der seinerseits einen Fehler auswirft,
    // erzeugt eine Schleife - und die faellt genau dann auf, wenn ohnehin schon etwas
    // nicht stimmt.
  } finally {
    reporting = false;
  }
}
