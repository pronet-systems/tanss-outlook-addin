/**
 * HTTP-Zugriff auf die TANSS-API - die einzige Stelle, an der dieses Add-in mit TANSS
 * spricht.
 *
 * Zwei Pfade mit gegensaetzlichen Regeln, und beide sind hier durchgesetzt, nicht
 * verabredet:
 *
 * `/api/v1/**`
 *     Braucht IMMER `loggedInUserId`. Ohne den Parameter antwortet die Route mit 403.
 *     Die Rechtepruefung greift dann als der genannte Mitarbeiter - es wird keine
 *     Berechtigung umgangen, sondern genau dessen benutzt.
 *
 * `/api/tanss.x/v1/**`
 *     Die Integrationsflaeche. Hier darf `loggedInUserId` NIEMALS anhaengen - nicht
 *     weil der Aufruf schiefginge, sondern weil der Server mit dem Parameter andere
 *     Zweige nimmt als die erprobten. Der Unterschied faellt erst im Betrieb auf.
 *
 * Der Kopf heisst `apiToken`, NICHT `Authorization`. Ein Token in `Authorization`
 * uebergeht TANSS stillschweigend, und der Aufruf scheitert danach mit einer Meldung,
 * die nach fehlenden Rechten aussieht.
 *
 * Status VOR Koerper. Ein leerer Koerper ist nur bei 2xx eine leere Antwort; bei 401,
 * 403 oder 502 ist er ein Fehler ohne Text und darf nie als "nichts gefunden"
 * durchgehen. Eine leere 403 saehe sonst aus wie "keine Treffer", und die
 * Terminaufloesung boete "Einsatz anlegen" an, obwohl laengst einer existiert.
 */

import { ApiError } from "./errors.js";

export const TANSS_X_PREFIX = "/api/tanss.x/v1";
export const V1_PREFIX = "/api/v1";

/**
 * Die einzige /api/v1-Route, die ohne Handelnden aufgerufen werden DARF: Sie stellt die
 * Identitaet ja erst her. Jeder weitere Eintrag waere ein Loch in der Grenze und
 * braeuchte eine Begruendung an dieser Stelle.
 */
export const ACTOR_FREE_V1_ROUTES = new Set(["/api/v1/login"]);

/** Zeitgrenze gewoehnlicher Aufrufe. Ohne sie haengt das Pane an einem stillen Proxy. */
export const DEFAULT_TIMEOUT_MS = 30_000;

/** Zeitgrenze fuer Uebertragungen mit Anhang - 25 MiB ueber eine Hotelleitung dauern. */
export const UPLOAD_TIMEOUT_MS = 180_000;

/** Statuswerte, bei denen ein zweiter Versuch ueberhaupt Sinn ergibt. */
const RETRYABLE_STATUS = new Set([429, 500, 502, 503, 504]);

/** Drei Versuche mit kurzem Rueckzug. Hinter dem Pane wartet ein Mensch. */
const RETRY_ATTEMPTS = 3;
const RETRY_BASE_DELAY_MS = 500;
const RETRY_MAX_DELAY_MS = 5_000;

/**
 * Fehlernamen aus TANSS auf den geschlossenen Code-Bestand der Oberflaeche abgebildet.
 *
 * Die Reihenfolge ist Absicht: erst die Namen, zuletzt der Status. TANSS verpackt sehr
 * unterschiedliche Sachverhalte in denselben Statuscode - ein Lizenzproblem wird zu
 * HTTP 500, eine abgelehnte Anmeldung zu HTTP 404 -, deshalb entscheidet der Text, wo
 * es einen gibt.
 */
const ERROR_NAMES = [
  ["INVALID_EML", "MAIL_UNPARSABLE"],
  ["CHANGES_WERE_DISCARDED", "APPOINTMENT_READONLY"],
  ["COMPANY_MUST_BE_GIVEN", "CONFLICT"],
  ["DUPLICATE", "CONFLICT"],
  ["FORBIDDEN_NO_COMPANY_ACCESS", "FORBIDDEN_NO_COMPANY_ACCESS"],
];

/** Statuszuordnung fuer Antworten, die keinen auswertbaren Fehlernamen tragen. */
const STATUS_CODES = new Map([
  [401, "UNAUTHENTICATED"],
  [403, "UNAUTHENTICATED"],
  [404, "CONFLICT"],
  [413, "MAIL_TOO_LARGE"],
  [429, "RATE_LIMITED"],
]);

/* ------------------------------------------------------------------ Hilfsmittel */

/** Wartet - die einzige Stelle mit einer Pause. */
function sleep(ms) {
  return new Promise((resolve) => globalThis.setTimeout(resolve, ms));
}

/**
 * Verbindet die Abbruchanforderung des Aufrufers mit der eigenen Zeitgrenze.
 * `AbortSignal.any` ist noch nicht ueberall vorhanden; die Handarbeit ist kurz genug.
 */
function withTimeout(signal, timeoutMs) {
  const controller = new AbortController();
  const timer = globalThis.setTimeout(() => controller.abort(new Error("timeout")), timeoutMs);
  if (signal) {
    if (signal.aborted) controller.abort(signal.reason);
    else signal.addEventListener("abort", () => controller.abort(signal.reason), { once: true });
  }
  return { signal: controller.signal, done: () => globalThis.clearTimeout(timer) };
}

/** `Retry-After` in Sekunden, sofern die Antwort einen brauchbaren Wert traegt. */
function retryAfterSeconds(response) {
  const raw = response.headers.get("Retry-After");
  if (!raw) return 0;
  const seconds = Number.parseInt(raw, 10);
  return Number.isFinite(seconds) && seconds > 0 ? Math.min(seconds, 60) : 0;
}

/**
 * Baut die Zieladresse aus Basis, Pfad und Abfrageparametern.
 *
 * Werte gehen durch `URLSearchParams`. Von Hand verkettet braeche jede Adresse mit
 * einem `+` oder `&` darin, und zwar still.
 */
function buildUrl(baseUrl, path, params) {
  // Die Basis darf relativ sein (`/backend`) - dann traegt das Dokument den Ursprung.
  // Ist sie absolut, bleibt das zweite Argument wirkungslos; genau das erlaubt es,
  // diese Funktion ausserhalb eines Browsers zu pruefen.
  const documentBase = globalThis.document ? globalThis.document.baseURI : undefined;
  const url = new URL(`${baseUrl}${path}`, documentBase);
  for (const [key, value] of Object.entries(params || {})) {
    if (value === null || value === undefined || value === "") continue;
    if (Array.isArray(value)) {
      for (const single of value) url.searchParams.append(key, String(single));
    } else {
      url.searchParams.set(key, String(value));
    }
  }
  return url.toString();
}

/**
 * Liest den Fehlernamen aus einer TANSS-Fehlerantwort.
 *
 * `error.text` traegt den unuebersetzten NAMEN des Falles, `error.localizedText`
 * denselben Fall als uebersetzten Satz. Vorrang hat der Name: Er ist stabil, der Satz
 * haengt an der Spracheinstellung der Instanz.
 */
function readError(payload) {
  if (!payload || typeof payload !== "object") {
    return { detail: "", tokenException: "", licenseError: "" };
  }
  const error = payload.error && typeof payload.error === "object" ? payload.error : {};
  return {
    detail: String(error.text || error.localizedText || payload.detail || ""),
    tokenException: String(error.tokenExceptionType || payload.tokenExceptionType || ""),
    licenseError: String(error.licenseError || ""),
  };
}

/**
 * Uebersetzt eine TANSS-Fehlerantwort in genau einen Code der Oberflaeche.
 *
 * Ein Lizenzproblem macht aus JEDER authentifizierten Anfrage einen HTTP 500. Ohne die
 * eigene Unterscheidung suchte der Techniker den Fehler an einer Stelle, an der keiner
 * ist - und ein zweiter Versuch hilft nie.
 */
export function codeFor(status, payload) {
  const { detail, tokenException, licenseError } = readError(payload);
  const upper = detail.toUpperCase();

  for (const [name, code] of ERROR_NAMES) {
    if (upper.includes(name)) return { code, detail };
  }
  if (licenseError || upper.includes("LICENSE") || upper.includes("LIZENZ")) {
    return { code: "TANSS_LICENSE_ERROR", detail: detail || licenseError };
  }
  if (tokenException.toUpperCase() === "EXPIRED" || upper.includes("TOKEN_HAS_EXPIRED")) {
    return { code: "UNAUTHENTICATED", detail };
  }
  if (status === 404 || upper.includes("OBJECT_NOT_FOUND")) {
    return { code: "CONFLICT", detail };
  }
  if (RETRYABLE_STATUS.has(status)) return { code: "TANSS_UNAVAILABLE", detail };
  return { code: STATUS_CODES.get(status) || "INTERNAL", detail };
}

/* ------------------------------------------------------------------ Der Client */

export class TanssClient {
  /**
   * @param {{baseUrl: string, tokens: {apiToken: () => string},
   *          employeeId?: () => number}} options
   *        `baseUrl` ist der Pfad, unter dem die TANSS-API liegt, ohne Schraegstrich
   *        am Ende. `tokens.apiToken()` liefert das aktuelle Token samt `Bearer `.
   *        `employeeId()` liefert den Handelnden - die EINZIGE Quelle des
   *        `loggedInUserId`.
   */
  constructor({ baseUrl, tokens, employeeId = () => 0 }) {
    this.baseUrl = String(baseUrl || "").replace(/\/+$/, "");
    this.tokens = tokens;
    this.employeeId = employeeId;
  }

  get(path, options = {}) {
    return this.call("GET", path, { retry: true, ...options });
  }

  put(path, json, options = {}) {
    return this.call("PUT", path, { json, ...options });
  }

  post(path, json, options = {}) {
    return this.call("POST", path, { json, ...options });
  }

  /**
   * Ein Aufruf gegen TANSS.
   *
   * @param {string} method
   * @param {string} path Mit fuehrendem Schraegstrich, z. B. `/api/v1/tickets`.
   * @param {{query?: object, json?: any, body?: BodyInit, headers?: object,
   *          retry?: boolean, wantMeta?: boolean, signal?: AbortSignal,
   *          timeoutMs?: number, actorFree?: boolean}} [options]
   *        `actorFree` unterdrueckt `loggedInUserId` fuer die Anmeldung selbst.
   * @returns {Promise<any>} Der `content` der Antwort, bei `wantMeta` ein Paar
   *          `{content, meta}`. Wirft `ApiError` mit einem Code der Oberflaeche.
   */
  async call(method, path, options = {}) {
    const params = { ...(options.query || {}) };

    if ("loggedInUserId" in params) {
      // Der Handelnde wird NIE aus einem uebergebenen Parameter abgeleitet. Traefe
      // dieser Fall ein, haette ein Aufrufer die einzige Stelle umgangen, an der er
      // entsteht - das ist ein Programmfehler, kein Datenfehler.
      throw new ApiError(
        "INTERNAL",
        "loggedInUserId darf nicht als Parameter uebergeben werden.",
      );
    }

    if (path.startsWith(TANSS_X_PREFIX)) {
      // Nichts anhaengen. Siehe Kopf dieser Datei.
    } else if (path.startsWith(V1_PREFIX)) {
      const actorFree = options.actorFree === true || ACTOR_FREE_V1_ROUTES.has(path);
      if (!actorFree) {
        const actor = Number(this.employeeId()) || 0;
        if (actor <= 0) {
          throw new ApiError(
            "UNAUTHENTICATED",
            "Es ist kein Mitarbeiter angemeldet. Ohne ihn darf nicht gehandelt werden.",
          );
        }
        params.loggedInUserId = actor;
      }
    }

    const url = buildUrl(this.baseUrl, path, params);
    const retry = options.retry === true;
    const timeoutMs = options.timeoutMs
      || (options.body ? UPLOAD_TIMEOUT_MS : DEFAULT_TIMEOUT_MS);

    let lastError = null;
    for (let attempt = 1; attempt <= (retry ? RETRY_ATTEMPTS : 1); attempt += 1) {
      let response;
      const guard = withTimeout(options.signal, timeoutMs);
      try {
        response = await fetch(url, {
          method,
          headers: this._headers(options),
          body: this._body(options),
          // Keine Cookies: TANSS nimmt ausschliesslich den `apiToken`-Kopf entgegen.
          // Damit ist eine Vorlage ueber eine fremde Seite strukturell ausgeschlossen,
          // nicht nur abgesichert.
          credentials: "omit",
          cache: "no-store",
          redirect: "error",
          signal: guard.signal,
        });
      } catch (error) {
        if (options.signal && options.signal.aborted) throw new ApiError("CLIENT_UNEXPECTED");
        lastError = new ApiError("CLIENT_NETWORK", "", { cause: error });
        if (!retry || attempt === RETRY_ATTEMPTS) throw lastError;
        await sleep(Math.min(RETRY_BASE_DELAY_MS * 2 ** (attempt - 1), RETRY_MAX_DELAY_MS));
        continue;
      } finally {
        guard.done();
      }

      if (retry && RETRYABLE_STATUS.has(response.status) && attempt < RETRY_ATTEMPTS) {
        // Erst in den Koerper sehen, dann ueber die Wiederholung entscheiden: TANSS
        // kleidet dauerhafte Zustaende in HTTP 500. Ohne diesen Blick verbraucht das
        // Pane drei Versuche und meldet am Ende "nicht erreichbar", waehrend in
        // Wahrheit der Wartungsvertrag abgelaufen ist.
        const peeked = await this._peek(response);
        if (peeked && (peeked.code === "TANSS_LICENSE_ERROR" || peeked.code === "UNAUTHENTICATED")) {
          throw new ApiError(peeked.code, "", { status: response.status, detail: peeked.detail });
        }
        const wait = retryAfterSeconds(response) * 1000
          || Math.min(RETRY_BASE_DELAY_MS * 2 ** (attempt - 1), RETRY_MAX_DELAY_MS);
        await sleep(wait);
        continue;
      }

      return this._handle(response, options.wantMeta === true);
    }

    throw lastError || new ApiError("TANSS_UNAVAILABLE");
  }

  _headers(options) {
    const headers = { Accept: "application/json", ...(options.headers || {}) };
    if (!headers.apiToken) {
      const token = this.tokens.apiToken();
      if (token) headers.apiToken = token.startsWith("Bearer ") ? token : `Bearer ${token}`;
    }
    // Kein Content-Type bei FormData: Die Grenzmarke des Multipart setzt der Browser.
    if (options.json !== undefined) headers["Content-Type"] = "application/json";
    return headers;
  }

  _body(options) {
    if (options.body !== undefined) return options.body;
    if (options.json !== undefined) return JSON.stringify(options.json);
    return undefined;
  }

  /** Liest eine Fehlerantwort, ohne sie zu verbrauchen - fuer die Wiederholungsfrage. */
  async _peek(response) {
    try {
      const clone = response.clone();
      const text = await clone.text();
      if (!text) return null;
      return codeFor(response.status, JSON.parse(text));
    } catch {
      return null;
    }
  }

  async _handle(response, wantMeta) {
    const status = response.status;
    const text = await response.text().catch(() => "");

    // Ein leerer Koerper ist NUR bei Erfolg eine leere Antwort.
    if (status >= 200 && status < 300 && (status === 204 || text === "")) {
      return wantMeta ? { content: {}, meta: {} } : {};
    }

    let payload = null;
    try {
      payload = text ? JSON.parse(text) : null;
    } catch {
      payload = null;
    }

    if (status >= 200 && status < 300) {
      if (payload === null) {
        // 2xx, aber kein JSON: So sieht eine Anmeldeseite aus, die ein Proxy
        // dazwischengeschoben hat.
        throw new ApiError("CLIENT_BAD_RESPONSE", "", { status });
      }
      if (typeof payload === "object" && payload !== null && "content" in payload) {
        return wantMeta
          ? { content: payload.content, meta: payload.meta || {} }
          : payload.content;
      }
      return wantMeta ? { content: payload, meta: {} } : payload;
    }

    const { code, detail } = codeFor(status, payload);
    throw new ApiError(code, "", { status, detail });
  }
}
