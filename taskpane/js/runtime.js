/**
 * Was das Pane zur Laufzeit braucht, an einer Stelle zusammengesetzt.
 *
 * Es gibt keinen Dienst mehr, der diese Teile verdrahten koennte - also tut es das
 * Pane beim Start selbst: Konfiguration lesen, Sitzung anlegen, Client und Fachschicht
 * darauf setzen, oertlichen Speicher auf Installation und Mitarbeiter zuschneiden.
 *
 * Ein einziges Objekt, kein verteilter Zustand. Alles, was hier haengt, gehoert
 * zusammen und aendert sich zusammen: Meldet sich jemand an, wechselt die
 * Mitarbeiternummer - und damit sowohl der Handelnde in jedem TANSS-Aufruf als auch
 * der Bereich im oertlichen Speicher. Zwei Besitzer dieser Information liefen
 * unweigerlich auseinander.
 */

import { config, load as loadConfig } from "./config.js";
import { Store } from "./store.js";
import { TanssClient } from "./tanss/client.js";
import { TanssRepository } from "./tanss/repository.js";
import { Session } from "./tanss/session.js";

const state = {
  /** @type {object|null} */ config: null,
  /** @type {Session|null} */ session: null,
  /** @type {TanssClient|null} */ client: null,
  /** @type {TanssRepository|null} */ repository: null,
  /** @type {Store|null} */ store: null,
  /** @type {Promise<void>|null} */ starting: null,
};

/**
 * Liest die Konfiguration und baut die Schichten darauf.
 *
 * Ein zweiter Aufruf liefert denselben Vorgang. Zwei parallel gebaute Sitzungen
 * haetten jede ihren eigenen Zwischenspeicher, und die Erneuerung des Tokens liefe
 * doppelt - beide verbrauchten dasselbe Erneuerungstoken, und die zweite scheiterte.
 */
export function start() {
  if (state.starting) return state.starting;
  state.starting = (async () => {
    const settings = await loadConfig();
    state.config = settings;

    state.session = new Session({ baseUrl: settings.apiBase });
    state.client = new TanssClient({
      baseUrl: settings.apiBase,
      tokens: { apiToken: () => state.session.apiToken() },
      // Der Handelnde kommt AUSSCHLIESSLICH aus der Sitzung. Es gibt im ganzen Pane
      // keinen zweiten Weg, `loggedInUserId` zu setzen - der Client weist einen
      // uebergebenen Wert als Programmfehler ab.
      employeeId: () => state.session.employeeId(),
    });
    state.repository = new TanssRepository({
      client: state.client,
      config: settings,
    });
    rebuildStore();
  })();
  return state.starting;
}

/**
 * Schneidet den oertlichen Speicher neu zu.
 *
 * Muss nach jeder Anmeldung laufen: Der Bereich haengt an der Mitarbeiternummer, damit
 * an einem geteilten Rechner niemand die Vorbelegungen und zuletzt benutzten Tickets
 * eines Kollegen sieht.
 */
export function rebuildStore() {
  state.store = new Store({
    instance: state.config ? state.config.instance : "",
    employeeId: state.session ? state.session.employeeId() : 0,
  });
}

export function settings() {
  return state.config || config();
}

/** @returns {Session} */
export function session() {
  if (!state.session) throw new Error("runtime.start() wurde nicht abgewartet");
  return state.session;
}

/** @returns {TanssRepository} */
export function repository() {
  if (!state.repository) throw new Error("runtime.start() wurde nicht abgewartet");
  return state.repository;
}

/** @returns {Store} */
export function store() {
  if (!state.store) rebuildStore();
  return state.store;
}

/**
 * Sorgt dafuer, dass ein gueltiges Token vorliegt, bevor TANSS gerufen wird.
 *
 * Wirft `ApiError("UNAUTHENTICATED")`, wenn keine Sitzung besteht oder die Erneuerung
 * scheitert. Der Verteiler in `api.js` macht daraus den Umschlag, den die Seiten
 * erwarten; die Anmeldeseite ist die einzige, die ohne diesen Schritt auskommt.
 */
export async function ready() {
  await start();
  await session().ensureFresh();
}
