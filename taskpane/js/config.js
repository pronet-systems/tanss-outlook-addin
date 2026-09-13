/**
 * Die Einstellungen dieser Installation.
 *
 * Es gibt keinen Dienst mehr, in dem eine Konfiguration liegen koennte - also bekommt
 * der Client sie mit: `config.json` liegt neben `index.html`, und das Pane liest sie
 * beim Start. Der Administrator aendert eine Datei auf dem Webserver; beim naechsten
 * Oeffnen gilt sie fuer alle.
 *
 * Drei Eigenschaften dieser Datei, die man kennen muss:
 *
 * - **Sie ist oeffentlich lesbar.** Sie liegt unter derselben Adresse wie das Pane und
 *   wird unauthentifiziert abgerufen. Deshalb steht in ihr KEIN Geheimnis, und es kann
 *   auch keines hineingeraten: Der Techniker meldet sich selbst an, es gibt in diesem
 *   Aufbau kein ruhendes Geheimnis.
 * - **Sie darf fehlen.** Fehlt sie oder ist sie unlesbar, gelten die Vorgaben aus
 *   dieser Datei, und das Pane oeffnet. Die Diagnoseseite sagt dann, dass die Vorgaben
 *   greifen - eine Installation soll nicht daran scheitern, dass jemand eine geschweifte
 *   Klammer vergessen hat.
 * - **Sie wird nicht zwischengespeichert.** `cache: "no-store"`: Eine Aenderung soll
 *   beim naechsten Oeffnen gelten und nicht erst, wenn ein Zwischenspeicher es zulaesst.
 *
 * Unbekannte Schluessel werden ignoriert, nicht uebernommen. Ein Tippfehler bleibt damit
 * wirkungslos, statt eine Einstellung zu erfinden, auf die sich niemand verlassen kann.
 */

/**
 * Die Vorgaben. Sie sind so gewaehlt, dass eine Installation, die nichts einstellt,
 * arbeitsfaehig ist.
 */
const DEFAULTS = {
  /**
   * Wo die TANSS-API liegt.
   *
   * Relativ (`/backend`), wenn Pane und API auf demselben Namen liegen - dann entfaellt
   * CORS. Absolut (`https://tanss.example.de/backend`), wenn das Pane woanders liegt;
   * das ist der Regelfall, sobald es einen eigenen Namen oder einen gemeinsamen
   * oeffentlichen Hoster hat. Die TANSS-API laesst den fremden Ursprung zu, und der
   * Hoster spricht nie selbst mit TANSS - jeder Aufruf geht vom Browser des Technikers
   * aus. Deshalb darf die TANSS-Instanz durchaus nur im Firmennetz erreichbar sein.
   */
  apiBase: "/backend",

  /**
   * Welche Adressen als TANSS-Instanz zugelassen sind, als Liste von Urspruengen.
   *
   * Leer heisst: jede https-Adresse. Das ist fuer einen gemeinsam genutzten Hoster
   * gedacht, auf dem jeder Kunde seine eigene Instanz mitbringt. Wer nur die eigene
   * Instanz zulassen will, traegt sie hier ein - dann wird eine abweichende Angabe aus
   * der Adresszeile abgewiesen statt befolgt.
   */
  allowedApiOrigins: [],

  /**
   * Wo die TANSS-Oberflaeche liegt. Daraus entstehen die Verweise "in TANSS oeffnen".
   * Leer heisst: keine Verweise anbieten, statt auf eine geratene Adresse zu zeigen.
   */
  frontendBase: "",

  /**
   * Kennzeichen dieser Installation. Es haengt an jeder Terminkopplung, damit eine
   * Einsatznummer aus einer anderen Installation nicht als die eigene gelesen wird -
   * ein Termin, der aus einem Testsystem kopiert wurde, haengt sonst an einer Nummer,
   * die es hier gibt, aber an einem ganz anderen Einsatz.
   * Leer: Das Pane leitet es aus der TANSS-Adresse ab - siehe `deriveInstance`.
   */
  instance: "",

  /**
   * Die Entra-Registrierung fuer den Zugriff auf die Originalnachricht.
   *
   * Die Id kommt bevorzugt aus der Adresse des Panes - sie steht im Manifest, das jeder
   * Kunde selbst einspielt. Der Eintrag hier ist der Rueckfall fuer eine Installation
   * mit eigener Ablage.
   */
  entra: {
    /**
     * Ohne diese Id gibt es keinen Weg an die vollstaendige Originalnachricht: Das Pane
     * setzt die Mail dann aus Office-Bordmitteln zusammen. Das ergibt eine
     * Rekonstruktion, keine Kopie - Empfangskette und Signaturkoepfe fehlen.
     * Die Registrierung braucht KEIN Geheimnis; sie ist eine oeffentliche
     * Clientanwendung mit delegierten Berechtigungen.
     */
    clientId: "",
  },

  limits: {
    /** Groesste Nachricht, die abgelegt wird. Daran haengt auch die Vorwarnung. */
    maxEmlBytes: 25 * 1024 * 1024,
  },

  features: {
    /** Die Terminseite anbieten. */
    appointments: true,
    /**
     * Selbst Einsaetze anlegen. Nur in einer Installation OHNE
     * Kalendersynchronisation einschalten - sonst entstehen zwei Eintraege je Termin.
     */
    createSupport: false,
    /** Die Nachricht ueber Graph holen. Braucht `entra.clientId`. */
    graphMail: true,
    /** Den Termin ueber Graph lesen. Braucht `entra.clientId`. */
    graphCalendar: false,
  },

  defaults: {
    /** Vorgewaehlte Leistungsart, wenn die Liste nicht abrufbar ist. */
    supportTypeId: 0,
    /** Vorgewaehlter Einsatzort. */
    location: "OFFICE",
    /** Neue Einsaetze zunaechst als intern kennzeichnen. */
    internal: true,
    /** Vorgewaehlter Tickettyp. 0 heisst: keine Vorauswahl. */
    ticketTypeId: 0,
  },

  /**
   * Die Tickettypen dieser Installation, als `[{id, name}]`.
   *
   * Sie stehen in der Konfiguration und nicht in einer Abfrage, weil TANSS keine Route
   * fuehrt, die sie vollstaendig und ungefiltert liefert.
   *
   * Auf einer gemeinsam genutzten Ablage nuetzt dieser Schluessel allerdings nichts: Dort
   * ist die `config.json` fuer alle Kunden dieselbe und laesst ihn leer. Tickettypen sind
   * aber kundenspezifisch - sie nehmen deshalb denselben Weg wie die TANSS-Adresse und
   * die Anwendungs-Id, naemlich ueber die Seitenadresse im Manifest. Siehe
   * `applyTicketTypes`.
   *
   * Bleibt die Liste leer, zeigt die Ticketmaske kein Typfeld. Das ist zulaessig, denn
   * ein Ticket ohne Typ nimmt TANSS an.
   */
  ticketTypes: [],
};

/** Die geladene Konfiguration. Bis `load()` gelaufen ist, sind es die Vorgaben. */
let current = structuredClone(DEFAULTS);

/** Was beim Laden schiefging - fuer die Diagnoseseite, nicht fuer den Techniker. */
let loadNote = "noch nicht geladen";

/**
 * Uebernimmt nur die Schluessel, die es in den Vorgaben gibt, und nur mit dem Typ, den
 * die Vorgabe hat.
 *
 * Das ist die Sperre gegen den stillen Tippfehler: `"maxEmlByte"` erzeugt keinen neuen
 * Schluessel, und `"appointments": "ja"` macht aus einem Schalter keine Zeichenkette.
 * Was nicht passt, bleibt auf der Vorgabe - die Diagnoseseite nennt es.
 */
function merge(target, source, path, rejected) {
  if (!source || typeof source !== "object") return;
  for (const [key, value] of Object.entries(source)) {
    const here = path ? `${path}.${key}` : key;
    if (!Object.prototype.hasOwnProperty.call(target, key)) {
      rejected.push(here);
      continue;
    }
    const fallback = target[key];
    if (Array.isArray(fallback)) {
      if (Array.isArray(value)) target[key] = value;
      else rejected.push(here);
    } else if (fallback !== null && typeof fallback === "object") {
      merge(fallback, value, here, rejected);
    } else if (typeof value === typeof fallback) {
      target[key] = value;
    } else {
      rejected.push(here);
    }
  }
}

/**
 * Laedt `config.json` neben dem Pane. Wirft nie.
 *
 * @returns {Promise<object>} Die geltende Konfiguration.
 */
export async function load() {
  const merged = structuredClone(DEFAULTS);
  const rejected = [];
  try {
    const response = await fetch("config.json", {
      credentials: "omit",
      cache: "no-store",
      headers: { Accept: "application/json" },
    });
    if (!response.ok) {
      loadNote = `config.json nicht abrufbar (HTTP ${response.status}) - Vorgaben gelten`;
    } else {
      merge(merged, await response.json(), "", rejected);
      loadNote = rejected.length
        ? `geladen; unbekannt oder falscher Typ: ${rejected.join(", ")}`
        : "geladen";
    }
  } catch (error) {
    loadNote = `config.json nicht lesbar (${error && error.name}) - Vorgaben gelten`;
  }

  applyApiBase(merged);
  applyEntraClientId(merged);
  applyTicketTypes(merged);
  if (!merged.instance) merged.instance = deriveInstance(merged);
  current = merged;
  return current;
}

/**
 * Ein Kennzeichen fuer diese Installation, wenn keines eingestellt ist.
 *
 * Abgeleitet wird es aus der TANSS-Adresse, NICHT aus dem Hostnamen der Seite. Der
 * Unterschied faellt erst bei einer gemeinsam genutzten Ablage auf, dann aber hart:
 * Dort ist der Hostname der Seite fuer ALLE Kunden derselbe. Zwei TANSS-Instanzen auf
 * einem Arbeitsplatz - Produktiv- und Testsystem etwa - teilten sich damit ihren
 * oertlichen Bereich, und ein Termin traege die Einsatznummer der jeweils anderen
 * Installation.
 *
 * Die TANSS-Adresse dagegen unterscheidet genau das, was unterschieden werden muss.
 * Sie ist kein Geheimnis; sie steht in jeder Adresse, die der Techniker ohnehin sieht.
 */
function deriveInstance(config) {
  const base = config.apiBase || "";
  if (base && !base.startsWith("/")) {
    try {
      return new URL(base).host;
    } catch {
      // Unlesbare Adresse - dann traegt der Hostname der Seite.
    }
  }
  const host = globalThis.location ? String(globalThis.location.hostname || "") : "";
  return host || "unbekannt";
}

/** Schluessel der selbst eingetragenen TANSS-Adresse im oertlichen Speicher. */
const API_BASE_KEY = "tanss.addin.apiBase";

/**
 * Legt fest, gegen welche TANSS-Instanz gearbeitet wird. Drei Stufen, in dieser
 * Reihenfolge:
 *
 * 1. **Der Parameter `tanss` in der Adresse des Panes.** Er stammt aus dem Manifest,
 *    und das Manifest spielt der Administrator des Mandanten ein - die Angabe ist also
 *    verwaltet, nicht vom Benutzer erraten. Genau das erlaubt EINEN gemeinsamen Hoster
 *    fuer beliebig viele Kunden: Jeder bringt sein eigenes Manifest mit.
 * 2. **`apiBase` aus der config.json.** Der Regelfall fuer eine Installation, die
 *    ihren eigenen Hoster betreibt.
 * 3. **Die zuletzt selbst eingetragene Adresse.** Fuer den Fall, dass weder das eine
 *    noch das andere vorliegt; das Pane fragt dann einmal danach.
 *
 * Zugelassen sind nur https-Adressen, und nur solche aus `allowedApiOrigins`, sofern
 * die Liste gefuellt ist. Der Grund ist nicht Formalismus: Der Techniker tippt sein
 * TANSS-Kennwort in dieses Pane. Eine untergeschobene Adresse waere eine
 * Anmeldemaske, die das Kennwort woanders hin schickt. Deshalb wird die Adresse
 * geprueft - und deshalb zeigt die Anmeldeseite sie an, statt sie zu verschweigen.
 */
function applyApiBase(merged) {
  const fromQuery = readQueryApi();
  const candidates = [fromQuery, merged.apiBase, readStoredApiBase()];
  for (const candidate of candidates) {
    if (!candidate) continue;
    // Eine relative Angabe meint denselben Ursprung und ist damit immer zulaessig.
    if (candidate.startsWith("/")) {
      merged.apiBase = candidate.replace(/\/+$/, "");
      return;
    }
    if (isAllowed(candidate, merged.allowedApiOrigins)) {
      merged.apiBase = candidate.replace(/\/+$/, "");
      return;
    }
  }
  merged.apiBase = "";
}

/**
 * Die Anwendungs-Id fuer die Microsoft-Anmeldung.
 *
 * Sie nimmt denselben Weg wie die TANSS-Adresse: Vorrang hat der Parameter aus der
 * Adresse des Panes, denn der stammt aus dem Manifest, und das spielt der Administrator
 * des Mandanten ein. Auf einer gemeinsam genutzten Ablage ist die config.json fuer alle
 * Kunden dieselbe - eine dort eingetragene Id waere entweder falsch oder erzwaenge je
 * Kunde eine eigene Ablage.
 *
 * Geprueft wird nur die Form: Eine Id, die keine Kennung ist, wird verworfen. Mehr ist
 * hier nicht zu pruefen - eine falsche Id fuehrt dazu, dass der Anmeldedienst die
 * Anwendung nicht kennt, und das meldet er deutlich.
 */
function applyEntraClientId(merged) {
  const ausDerAdresse = readQueryParam("entra");
  const kandidat = ausDerAdresse || merged.entra.clientId || "";
  merged.entra.clientId = isClientId(kandidat) ? kandidat : "";
}

/**
 * Die Tickettypen aus der Seitenadresse - der einzige Weg, der auf einer gemeinsam
 * genutzten Ablage traegt.
 *
 * Form: `typen=1:Stoerung,2:Anfrage`. Kennung und Name durch Doppelpunkt getrennt,
 * Eintraege durch Komma. Der Name darf Doppelpunkte enthalten - getrennt wird am ERSTEN
 * -, aber kein Komma; der Generator laesst keines durch.
 *
 * Was nicht lesbar ist, wird still uebergangen und nicht zu einem halben Eintrag
 * gemacht: Ein Tickettyp ohne Kennung waere in der Auswahlliste sichtbar und beim
 * Anlegen wirkungslos - der Techniker waehlte etwas aus, das nichts tut.
 *
 * Die Adresse sticht die Datei. Beides gleichzeitig gibt es nur, wenn jemand eine eigene
 * Ablage betreibt UND ein Manifest mit Typen einspielt; dann gilt das Manifest, weil es
 * vom Administrator des Mandanten stammt und damit verwaltet ist.
 */
function applyTicketTypes(merged) {
  const roh = readQueryParam("typen");
  if (!roh) return;

  const gelesen = [];
  for (const stueck of roh.split(",")) {
    const trenner = stueck.indexOf(":");
    if (trenner < 1) continue;
    const id = Number(stueck.slice(0, trenner).trim());
    const name = stueck.slice(trenner + 1).trim();
    if (!Number.isInteger(id) || id <= 0 || name === "") continue;
    if (gelesen.some((eintrag) => eintrag.id === id)) continue;
    gelesen.push({ id, name });
  }

  if (gelesen.length > 0) merged.ticketTypes = gelesen;
}

/** Sieht das wie eine Anwendungs-Id aus? */
export function isClientId(value) {
  return /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/
    .test(String(value || "").trim());
}

function readQueryApi() {
  return readQueryParam("tanss");
}

/** Ein Parameter aus der Adresse des Panes. Leer, wenn es ihn nicht gibt. */
function readQueryParam(name) {
  try {
    const value = new URL(globalThis.location.href).searchParams.get(name);
    return value ? String(value).trim() : "";
  } catch {
    return "";
  }
}

function readStoredApiBase() {
  try {
    return String(globalThis.localStorage.getItem(API_BASE_KEY) || "");
  } catch {
    return "";
  }
}

/**
 * Ist diese Adresse als TANSS-Instanz zugelassen?
 *
 * Nur https - ueber http liefe das Kennwort im Klartext. Und nur ein Ursprung aus der
 * Liste, sofern es eine gibt; eine leere Liste laesst jede https-Adresse zu.
 */
export function isAllowed(url, allowed = current.allowedApiOrigins) {
  let parsed;
  try {
    parsed = new URL(url);
  } catch {
    return false;
  }
  if (parsed.protocol !== "https:") return false;
  if (!allowed || allowed.length === 0) return true;
  return allowed.some((entry) => {
    try {
      return new URL(entry).origin === parsed.origin;
    } catch {
      return false;
    }
  });
}

/**
 * Traegt eine selbst eingegebene TANSS-Adresse ein.
 *
 * @returns {boolean} Ob sie uebernommen wurde.
 */
export function setApiBase(url) {
  const value = String(url || "").trim().replace(/\/+$/, "");
  if (!isAllowed(value)) return false;
  current.apiBase = value;
  try {
    globalThis.localStorage.setItem(API_BASE_KEY, value);
  } catch {
    // Ohne Speicher gilt die Angabe nur fuer dieses Fenster - besser als gar nicht.
  }
  return true;
}

/**
 * Der Name der TANSS-Instanz, gegen die gearbeitet wird - fuer die Anzeige auf der
 * Anmeldeseite. Leer, wenn noch keine feststeht.
 */
export function apiHost() {
  const base = current.apiBase || "";
  if (!base) return "";
  if (base.startsWith("/")) {
    return globalThis.location ? String(globalThis.location.host || "") : "";
  }
  try {
    return new URL(base).host;
  } catch {
    return "";
  }
}

/** Die geltende Konfiguration. */
export function config() {
  return current;
}

/** Was beim Laden geschah - ausschliesslich fuer die Diagnoseseite. */
export function note() {
  return loadNote;
}

export { DEFAULTS };
