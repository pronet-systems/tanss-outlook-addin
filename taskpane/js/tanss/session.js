/**
 * Anmeldung, Erneuerung und Aufbewahrung des Tokens - der einzige Ort, an dem in dieser
 * Version ein Token entsteht.
 *
 * Es gibt keinen eigenen Dienst mehr, der im Namen des Technikers handeln koennte.
 * Deshalb meldet sich der Techniker selbst an und bekommt sein EIGENES Token. Der
 * Unterschied ist der Kern dieses Zuschnitts: Bei einem Anmeldetoken steht der
 * Mitarbeiter in der signierten Aussage des Tokens und ist nicht faelschbar. Ein
 * Anwendungstoken dagegen nennt den Mitarbeiter als Parameter - ein solches Token
 * gehoert nie in einen Browser.
 *
 * Laufzeiten, an denen dieses Modul haengt:
 *
 * - Das Zugriffstoken (`apiKey`) gilt VIER STUNDEN.
 * - Das Erneuerungstoken (`refresh`) gilt deutlich laenger und wird bei jeder
 *   Erneuerung mit erneuert.
 *
 * Erneuert wird ohne Zutun: Der Kopf `refreshToken` an einer beliebigen Route - ausser
 * der Anmeldung selbst - liefert ein NEUES PAAR statt der eigentlichen Antwort. Der
 * Aufruf, den man dafuer stellt, wird also nicht beantwortet; deshalb steht dafuer eine
 * eigene, billige Route und nicht der gerade anstehende Fachaufruf.
 *
 * Was dieser Kopf voraussetzt: Er ist ein EIGENER Kopf und loest damit eine
 * CORS-Vorabfrage aus. Liegt das Pane auf einem anderen Ursprung als die API - bei einer
 * gemeinsam genutzten Ablage ist das der Regelfall -, muss `refreshToken` in der
 * Antwort-Kopfzeile `Access-Control-Allow-Headers` der TANSS-Instanz stehen. Steht er
 * nicht darin, gelingen Anmeldung und alle Fachaufrufe (`apiToken` steht dort), aber die
 * Erneuerung scheitert AUSNAHMSLOS - und zwar erst nach vier Stunden, wenn das erste
 * Zugriffstoken ablaeuft. Zu pruefen ist das ohne Browser:
 *
 *   curl -si -X OPTIONS "<apiBase>/api/v1/employees/ownState" \
 *     -H "Origin: <ursprung-des-panes>" \
 *     -H "Access-Control-Request-Method: GET" \
 *     -H "Access-Control-Request-Headers: refreshtoken" | grep -i allow-headers
 *
 * Wo das liegt: `localStorage` des eigenen Ursprungs - und zwar bewusst, obwohl
 * `sessionStorage` weniger aufbewahrte. Ein Taskpane wird bei jedem Oeffnen neu
 * aufgebaut; mit `sessionStorage` stuende vor jedem Arbeitsschritt eine Anmeldung. Der
 * Ursprung ist derselbe, unter dem der Techniker die TANSS-Oberflaeche ohnehin
 * bedient - das Token gilt also fuer genau die Stelle, an der es liegt, und nicht fuer
 * eine weitere.
 */

import { ApiError } from "./errors.js";

/** Schluessel im Speicher. Der Ursprung trennt bereits; ein Praefix genuegt. */
const STORAGE_KEY = "tanss.addin.session";

/**
 * Vorlauf, mit dem ein Token als abgelaufen gilt. Deckt Uhrenversatz zwischen
 * Arbeitsplatz und Server sowie die Laufzeit des Aufrufs ab.
 */
const EXPIRY_SKEW_MS = 120_000;

/**
 * Die Route, an der erneuert wird.
 *
 * Sie muss zwei Dinge erfuellen: billig sein und NICHT die Anmeldung selbst. Diese hier
 * liefert einen kleinen Satz ueber den eigenen Zustand; wird der Kopf `refreshToken`
 * ausgewertet, kommt statt dessen das neue Tokenpaar zurueck. Genau daran erkennt
 * `renew()` auch, ob die Erneuerung ueberhaupt gegriffen hat.
 */
const RENEW_PATH = "/api/v1/employees/ownState";

/* ------------------------------------------------------------------ Speicher */

/**
 * Lesen und Schreiben, die nie werfen.
 *
 * Manche Hosts sperren den Speicher (privates Fenster, geleerte Seitendaten, eine
 * Vorschau beim Erzeugen eines Vorschaubildes). Dann traegt der Arbeitsspeicher die
 * Sitzung: Sie ueberlebt dann kein Schliessen des Panes, aber sie funktioniert.
 */
let memory = null;

function read() {
  if (memory) return memory;
  try {
    const raw = globalThis.localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    const record = JSON.parse(raw);
    return record && typeof record === "object" ? record : null;
  } catch {
    return null;
  }
}

function write(record) {
  memory = record;
  try {
    if (record === null) globalThis.localStorage.removeItem(STORAGE_KEY);
    else globalThis.localStorage.setItem(STORAGE_KEY, JSON.stringify(record));
  } catch {
    // Kein Speicher verfuegbar - der Arbeitsspeicher traegt die Sitzung.
  }
}

/* ------------------------------------------------------------------ Sitzung */

export class Session {
  /**
   * @param {{baseUrl: string, fetchImpl?: typeof fetch}} options
   *        `fetchImpl` existiert nur, damit sich Anmeldung und Erneuerung ohne Browser
   *        pruefen lassen.
   */
  constructor({ baseUrl, fetchImpl = null }) {
    this.baseUrl = String(baseUrl || "").replace(/\/+$/, "");
    this.fetchImpl = fetchImpl || ((...args) => globalThis.fetch(...args));
    /** @type {Promise<void>|null} Laufende Erneuerung - die Einfachfuehrung. */
    this._renewing = null;
  }

  /* ---------------------------------------------------------------- Auskunft */

  /** Die Mitarbeiternummer des Angemeldeten, oder 0. Quelle des `loggedInUserId`. */
  employeeId() {
    const record = read();
    return record && Number(record.employeeId) > 0 ? Number(record.employeeId) : 0;
  }

  /** Der Anzeigename, den die Anmeldung hergab. */
  username() {
    const record = read();
    return record ? String(record.username || "") : "";
  }

  /** Das Zugriffstoken samt `Bearer `-Praefix, oder eine leere Zeichenkette. */
  apiToken() {
    const record = read();
    return record ? String(record.apiToken || "") : "";
  }

  /** Liegt ein Zugriffstoken vor, das noch lange genug gilt? */
  isFresh() {
    const record = read();
    if (!record || !record.apiToken) return false;
    return Number(record.expiresAt || 0) - EXPIRY_SKEW_MS > Date.now();
  }

  /** Gibt es ueberhaupt eine Sitzung - notfalls eine, die erst erneuert werden muss? */
  exists() {
    const record = read();
    return Boolean(record && (record.apiToken || record.refreshToken));
  }

  /* ---------------------------------------------------------------- Anmeldung */

  /**
   * Meldet an TANSS an.
   *
   * Die Zugangsdaten stehen ausschliesslich im Koerper, nie in der Abfrage: Ein
   * Parameter in der Adresse landete in jedem Proxy-Protokoll, in jedem Zugriffslog und
   * im Verlauf des Browsers.
   *
   * @param {{username: string, password: string, token?: string}} credentials
   *        `token` ist die Einmalkennung der Zwei-Faktor-Anmeldung und darf fehlen.
   */
  async login({ username, password, token = "" }) {
    const body = { username, password };
    if (token) body.token = token;

    const response = await this._send("POST", "/api/v1/login", {
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });

    const content = await this._content(response);
    if (!response.ok || !content) {
      // Jede abgelehnte Anmeldung kommt ohne brauchbaren Grund zurueck - derselbe
      // Ausgang deckt falsches Kennwort, unbekannten Benutzer und eine fehlende oder
      // falsche Einmalkennung ab. Deshalb nennt die Meldung alle drei.
      throw new ApiError(
        "UNAUTHENTICATED",
        "Benutzername, Kennwort oder Einmalkennung stimmen nicht.",
        { status: response.status },
      );
    }

    this._store(content, username);
    return { employeeId: this.employeeId(), username: this.username() };
  }

  /**
   * Sorgt dafuer, dass ein gueltiges Zugriffstoken vorliegt.
   *
   * Mehrfachaufrufe teilen sich EIN Promise. Zwei parallele Erneuerungen verbrauchten
   * sonst beide dasselbe Erneuerungstoken, und die zweite scheiterte - mit dem
   * Ergebnis, dass der Techniker mitten in der Arbeit auf der Anmeldeseite landet.
   */
  async ensureFresh() {
    if (this.isFresh()) return;
    if (!this._renewing) {
      this._renewing = this.renew().finally(() => {
        this._renewing = null;
      });
    }
    await this._renewing;
  }

  /**
   * Holt mit dem Erneuerungstoken ein neues Paar.
   *
   * Gesendet wird NUR der Kopf `refreshToken`. Haengt zusaetzlich ein Zugriffstoken an,
   * beantwortet TANSS die Anfrage gewoehnlich und erneuert nichts - die Sitzung liefe
   * dann still ab, obwohl bei jedem Aufruf eine Erneuerung versucht wurde.
   */
  async renew() {
    const record = read();
    const refresh = record ? String(record.refreshToken || "") : "";
    if (!refresh) {
      throw new ApiError("UNAUTHENTICATED", "Es liegt keine Anmeldung vor.");
    }

    let response;
    try {
      response = await this._send("GET", RENEW_PATH, {
        headers: { refreshToken: refresh },
        query: { loggedInUserId: this.employeeId() },
      });
    } catch (error) {
      // `renew` als Kennzeichen, und zwar aus einem Grund, der im Betrieb teuer war:
      // Ein Netzfehler AN DIESER STELLE sieht aus wie jeder andere, hat aber eine
      // andere Behandlung. Ein gescheiterter Fachaufruf wird spaeter wieder gehen; eine
      // gescheiterte Erneuerung dagegen kann dauerhaft unmoeglich sein - etwa wenn die
      // Gegenstelle den Kopf `refreshToken` nicht durch die CORS-Vorabfrage laesst
      // (`Access-Control-Allow-Headers`). Dann scheitert JEDER weitere Versuch mit
      // demselben Fehler, und ohne dieses Kennzeichen bliebe der Oberflaeche nur der
      // Knopf "Erneut versuchen", der nie etwas anderes bewirken kann. Mit ihm kann sie
      // stattdessen den einzigen Ausweg anbieten: eine neue Anmeldung.
      throw new ApiError("CLIENT_NETWORK", "", { cause: error, detail: "renew" });
    }

    const content = response.ok ? await this._content(response) : null;
    if (!content || !content.apiKey) {
      // Entweder ist das Erneuerungstoken abgelaufen, oder die Antwort ist die
      // gewoehnliche der Route - dann wurde der Kopf nicht ausgewertet. Beides heisst
      // dasselbe: Ab hier hilft nur eine neue Anmeldung.
      this.clear();
      throw new ApiError(
        "UNAUTHENTICATED",
        "Die Anmeldung ist abgelaufen. Bitte neu anmelden.",
        { status: response.status },
      );
    }
    this._store(content, this.username());
  }

  /**
   * Was ueber die Sitzung gesagt werden darf - fuer die Diagnoseseite.
   *
   * Ausdruecklich OHNE Token und ohne Auszug davon. Ein Tokenanfang in einer Meldung, die
   * weitergereicht wird, ist ein Token in fremder Hand; und fuer die Frage, die hier zu
   * beantworten ist, braucht es ihn nicht.
   *
   * Diese Frage lautet: Warum steht nach einem Neustart die Anmeldemaske da? Zwei
   * Ursachen sehen von aussen gleich aus - der Speicher war leer, oder das
   * Erneuerungstoken wurde abgewiesen -, verlangen aber voellig verschiedene
   * Reparaturen. `hasRefresh` und `expiresAt` trennen sie.
   */
  describe() {
    const record = read();
    if (!record) return { stored: false, hasToken: false, hasRefresh: false, expiresAt: 0 };
    return {
      stored: true,
      hasToken: Boolean(record.apiToken),
      hasRefresh: Boolean(record.refreshToken),
      expiresAt: Number(record.expiresAt || 0),
      employeeId: Number(record.employeeId || 0),
      username: String(record.username || ""),
    };
  }

  /** Verwirft die Sitzung vollstaendig. */
  clear() {
    write(null);
    memory = null;
  }

  /* ---------------------------------------------------------------- intern */

  /**
   * Uebernimmt ein Tokenpaar.
   *
   * `expire` ist ein Zeitstempel in SEKUNDEN seit 1970. Wird er als Millisekunden
   * gelesen, gilt das Token rechnerisch bis 1970 und jede Anfrage erneuerte vorher -
   * deshalb steht die Umrechnung hier an genau einer Stelle.
   */
  _store(content, username) {
    const employeeId = Number(content.employeeId) || 0;
    if (employeeId <= 0) {
      throw new ApiError(
        "UNAUTHENTICATED",
        "TANSS hat die Anmeldung angenommen, aber keine Mitarbeiternummer geliefert.",
      );
    }
    const expireSeconds = Number(content.expire) || 0;
    write({
      employeeId,
      username: String(username || ""),
      apiToken: String(content.apiKey || ""),
      refreshToken: String(content.refresh || ""),
      // Ohne brauchbare Angabe lieber kurz halten als lange falsch liegen: vier
      // Stunden sind die zugesagte Laufzeit.
      expiresAt: expireSeconds > 0 ? expireSeconds * 1000 : Date.now() + 4 * 3_600_000,
    });
  }

  async _send(method, path, { headers = {}, body, query } = {}) {
    const documentBase = globalThis.document ? globalThis.document.baseURI : undefined;
    const url = new URL(`${this.baseUrl}${path}`, documentBase);
    for (const [key, value] of Object.entries(query || {})) {
      if (value !== null && value !== undefined && value !== "") {
        url.searchParams.set(key, String(value));
      }
    }
    return this.fetchImpl(url.toString(), {
      method,
      headers: { Accept: "application/json", ...headers },
      body,
      credentials: "omit",
      cache: "no-store",
      redirect: "error",
    });
  }

  /** Der `content` einer Antwort, oder `null`, wenn keiner lesbar ist. */
  async _content(response) {
    try {
      const text = await response.text();
      if (!text) return null;
      const payload = JSON.parse(text);
      if (payload && typeof payload === "object" && "content" in payload) {
        return payload.content;
      }
      return payload;
    } catch {
      return null;
    }
  }
}
