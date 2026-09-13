/**
 * Die Microsoft-Anmeldung - und nur noch sie.
 *
 * Frueher stand hier der Hauptweg: ein Access-Token auf die eigene Dienst-Ressource,
 * mit dem sich das Pane bei seinem Backend auswies. Das Backend gibt es nicht mehr. Der
 * Techniker meldet sich stattdessen selbst an TANSS an; sein Anmeldetoken traegt den
 * Mitarbeiter in der signierten Aussage und ist damit nicht faelschbar. Die ganze
 * Ressourcen- und Zustimmungsmechanik ist damit entfallen.
 *
 * Geblieben ist genau EIN Zweck: ein Token fuer Microsoft Graph, mit dem das Pane die
 * geoeffnete Nachricht im ORIGINALFORMAT holt. Ohne dieses Token setzt `mime.js` die
 * Nachricht aus Office-Bordmitteln zusammen - das ergibt eine Rekonstruktion, keine
 * Kopie: Empfangskette und Signaturkoepfe fehlen dann.
 *
 * Drei Festlegungen, die nicht verhandelbar sind:
 *
 * 1. **Dieses Token erreicht TANSS nie.** Es existiert ausschliesslich im Browser des
 *    angemeldeten Technikers und geht ausschliesslich an Microsoft. Daran haengt die
 *    Zusage, dass dieses Add-in keine Reichweite auf fremde Postfaecher hat.
 * 2. **Nichts wird dauerhaft gespeichert.** Der Zwischenspeicher liegt im
 *    `sessionStorage` und endet mit dem Fenster.
 * 3. **Die Sperre gegen Mehrfachanforderung ist ein echtes Promise.** Ein Schalter mit
 *    kurzer Nachfrage liefert unter Last zwei parallele Anforderungen und damit zwei
 *    Anmeldedialoge.
 */

import { T } from "../i18n/de.js";
import { settings, start } from "./runtime.js";
import { session } from "./runtime.js";
import { getUserProfile } from "./office.js";

/** Fehler mit einem Code aus `T.errors`. Alles, was hier herausfaellt, ist anzeigbar. */
export class AuthError extends Error {
  constructor(code, cause = null) {
    super(code);
    this.name = "AuthError";
    this.code = code;
    this.cause = cause;
  }
}

export const AUTH_NAA = "naa";
export const AUTH_TANSS = "tanss_login";
export const AUTH_NONE = "none";

/** Basis aller Graph-Bereiche. Steht hier, weil `mime.js` nur den kurzen Namen kennt. */
const GRAPH_SCOPE_BASE = "https://graph.microsoft.com/";

/** Vorlauf, mit dem ein Token als abgelaufen gilt. Deckt Uhrenversatz und Laufzeit ab. */
const EXPIRY_SKEW_MS = 120_000;

const state = {
  /** @type {Promise<void>|null} */ initPromise: null,
  /** @type {object|null} */ pca: null,
  clientId: "",
  authority: "",
  redirectUri: "",
  /** @type {boolean} Ob wirklich die Bruecke des Wirts vorliegt - siehe `hasNaaBridge`. */
  naaBridge: false,
  accountUsername: "",
  /** @type {string} Letzter Fehler im Klartext - nur fuer die Diagnoseseite. */
  lastError: "",
};

/** Zwischenspeicher je Bereich: `{token, expiresAt}`. Nur im Arbeitsspeicher. */
const tokenCache = new Map();

/** Laufende Anforderungen je Bereich. Das ist die Einfachfuehrung - ein echtes Promise. */
const inflight = new Map();

/* ------------------------------------------------------------------ Einrichtung */

/**
 * Baut die Clientanwendung, falls eine Anwendungs-Id hinterlegt ist.
 *
 * Die Id kommt aus der `config.json` und nicht mehr aus dem Manifest: Das Manifest ist
 * bei einem gemeinsam genutzten Hoster fuer alle Kunden dasselbe, die Registrierung
 * dagegen gehoert je Mandant.
 *
 * Ein zweiter Aufruf liefert denselben Vorgang - zwei Instanzen auf derselben
 * Anwendungs-Id vertraegt die Bibliothek nicht, und die Ereignisbindungen blieben
 * doppelt.
 */
export function initialize() {
  if (state.initPromise) return state.initPromise;
  state.initPromise = (async () => {
    try {
      await start();
      const config = settings();
      const clientId = String(config.entra.clientId || "").trim();
      if (!clientId) {
        // Keine Registrierung hinterlegt: Das Pane arbeitet mit Bordmitteln weiter.
        return;
      }
      state.clientId = clientId;

      // Der Wirt reicht den bereits angemeldeten Benutzer durch; die Rueckadresse
      // benennt dabei den Office-Wirt und nicht eine Webseite.
      state.redirectUri = `brk-multihub://${globalThis.location.host}`;
      state.authority = "https://login.microsoftonline.com/common";

      const msal = globalThis.msal;
      if (!msal || typeof msal.createNestablePublicClientApplication !== "function") {
        state.lastError = "msal-browser nicht geladen";
        return;
      }

      state.pca = await msal.createNestablePublicClientApplication({
        auth: {
          clientId,
          authority: state.authority,
          redirectUri: state.redirectUri,
          supportsNestedAppAuth: true,
        },
        cache: {
          // sessionStorage statt localStorage: Nichts soll das Schliessen des Panes
          // ueberleben. Das TANSS-Token lebt laenger, aber es ist auch das einzige,
          // das der Techniker selbst erneuern kann.
          cacheLocation: "sessionStorage",
          storeAuthStateInCookie: false,
        },
        system: {
          // Keine Telemetrie an Dritte - auch nicht die der Bibliothek.
          loggerOptions: { loggerCallback: () => {}, piiLoggingEnabled: false },
        },
      });

      state.naaBridge = hasNaaBridge(state.pca);

      const accounts = state.pca.getAllAccounts();
      if (accounts.length > 0) {
        state.pca.setActiveAccount(accounts[0]);
        state.accountUsername = accounts[0].username || "";
      }
    } catch (error) {
      state.lastError = describe(error);
    }
  })();
  return state.initPromise;
}

/**
 * Ob die erzeugte Clientanwendung wirklich ueber die Bruecke des Wirts arbeitet.
 *
 * Die Erzeugung faellt STILL auf eine gewoehnliche Clientanwendung zurueck, wenn der
 * Wirt keine Bruecke stellt. Ungeprueft uebernommen hiesse das zweierlei: Die
 * Diagnoseseite meldete einen Weg, auf dem nichts stattgefunden hat, und die Rueckfrage
 * schickte den Browser an eine Adresse, zu der kein Browser zurueckfindet - das
 * Anmeldefenster schloesse dann nie ab.
 */
function hasNaaBridge(pca) {
  const controller = pca ? pca.controller : null;
  if (controller && typeof controller === "object") return Boolean(controller.bridgeProxy);
  return typeof globalThis.nestedAppAuthBridge !== "undefined";
}

/* ----------------------------------------------------------------- Tokenbezug */

/**
 * Access-Token fuer Microsoft Graph, etwa `getGraphToken("Mail.Read")`.
 *
 * Dieses Token erreicht TANSS NIE - es existiert ausschliesslich im Browser des
 * angemeldeten Technikers und geht ausschliesslich an Microsoft.
 */
export function getGraphToken(shortScope) {
  const key = `graph:${shortScope}`;
  const cached = tokenCache.get(key);
  if (cached && cached.expiresAt - EXPIRY_SKEW_MS > Date.now()) {
    return Promise.resolve(cached.token);
  }

  // Einfachfuehrung: Wer waehrend einer laufenden Anforderung fragt, bekommt DASSELBE
  // Promise. Zwei parallele Anforderungen kosteten im besten Fall zwei Broker-Runden
  // und oeffneten im schlechtesten zwei Anmeldedialoge.
  const running = inflight.get(key);
  if (running) return running;

  const promise = (async () => {
    await initialize();
    // Ohne Bruecke gibt es keinen Weg zu einem Graph-Token: Die Rueckfrage einer
    // gewoehnlichen Clientanwendung liefe auf eine Adresse, von der sie nie zurueckkaeme.
    if (!state.pca || !state.naaBridge) throw new AuthError("CLIENT_AUTH_UNAVAILABLE");
    const result = await acquire([`${GRAPH_SCOPE_BASE}${shortScope}`]);
    tokenCache.set(key, result);
    return result.token;
  })().finally(() => {
    inflight.delete(key);
  });

  inflight.set(key, promise);
  return promise;
}

/**
 * Still, bei Bedarf mit Rueckfrage.
 *
 * Eine Richtlinie fuer bedingten Zugriff kann mitten im Betrieb einen neuen Nachweis
 * verlangen und schickt die Anforderung im Fehler mit. Wer sie ignoriert, versucht
 * still weiter und scheitert fuer immer - die Rueckfrage MUSS sie unveraendert
 * mitfuehren, sonst stellt der Anmeldedienst dasselbe Token noch einmal aus.
 */
async function acquire(scopes) {
  const account = await ensureAccount(scopes);
  try {
    const result = await state.pca.acquireTokenSilent({
      scopes,
      account: account || undefined,
    });
    remember(result);
    return toResult(result);
  } catch (error) {
    const claims = extractClaims(error);
    if (!claims && !isInteractionRequired(error)) throw toAuthError(error);
    try {
      const result = await state.pca.acquireTokenPopup({
        scopes,
        account: account || undefined,
        claims: claims || undefined,
      });
      remember(result);
      return toResult(result);
    } catch (popupError) {
      throw toAuthError(popupError);
    }
  }
}

/**
 * Ein Konto fuer die stille Anforderung.
 *
 * Der Wirt reicht es in aller Regel durch; fehlt es doch, hilft die Postfachadresse als
 * HINWEIS weiter. Der Hinweis waehlt nur das Konto vor - Nachweis ist er nie.
 */
async function ensureAccount(scopes) {
  const active = state.pca.getActiveAccount();
  if (active) return active;

  const accounts = state.pca.getAllAccounts();
  if (accounts.length > 0) {
    state.pca.setActiveAccount(accounts[0]);
    state.accountUsername = accounts[0].username || "";
    return accounts[0];
  }

  const profile = getUserProfile();
  const loginHint = profile && profile.emailAddress ? profile.emailAddress : "";
  if (loginHint && typeof state.pca.ssoSilent === "function") {
    try {
      const result = await state.pca.ssoSilent({ scopes, loginHint });
      remember(result);
      return result.account || null;
    } catch (error) {
      state.lastError = describe(error);
    }
  }
  return null;
}

function remember(result) {
  if (result && result.account) {
    state.pca.setActiveAccount(result.account);
    state.accountUsername = result.account.username || "";
  }
}

function toResult(result) {
  if (!result || typeof result.accessToken !== "string" || result.accessToken === "") {
    throw new AuthError("CLIENT_AUTH_FAILED");
  }
  const expiresOn = result.expiresOn instanceof Date ? result.expiresOn.getTime() : 0;
  return {
    token: result.accessToken,
    // Ohne Ablaufangabe lieber kurz halten als lange falsch liegen.
    expiresAt: expiresOn > 0 ? expiresOn : Date.now() + 5 * 60_000,
  };
}

/* --------------------------------------------------------- TANSS-Anmeldung */

/**
 * Besteht eine TANSS-Anmeldung?
 *
 * Sie ist in dieser Version der Regelweg und nicht mehr der Rueckfall - der Name
 * stammt aus der Zeit, als sich das Pane bei einem eigenen Dienst auswies.
 */
export function hasFallbackSession() {
  try {
    return session().exists();
  } catch {
    return false;
  }
}

/** Meldet ab. */
export function clearFallbackSession() {
  try {
    session().clear();
  } catch {
    // Die Laufzeit stand noch nicht - dann gibt es auch nichts zu raeumen.
  }
  tokenCache.clear();
}

/** Verwirft alle Microsoft-Token im Arbeitsspeicher. */
export function forgetTokens() {
  tokenCache.clear();
}

/* ------------------------------------------------------------------- Diagnose */

/** Ob eine Entra-Anwendung konfiguriert ist - entscheidet ueber den Graph-Weg. */
export function hasEntraApp() {
  return state.pca !== null;
}

/** Name des benutzten Anmeldeweges, als fertiger Anzeigetext. */
export function methodLabel() {
  if (hasFallbackSession()) return T.auth.methodTanss;
  return T.auth.methodNone;
}

export function method() {
  return hasFallbackSession() ? AUTH_TANSS : AUTH_NONE;
}

/** Rohwerte fuer die Diagnoseseite. Enthaelt bewusst kein Token und keinen Auszug. */
export function diagnostics() {
  let employeeId = 0;
  let username = "";
  try {
    employeeId = session().employeeId();
    username = session().username();
  } catch {
    // Die Laufzeit stand noch nicht.
  }
  return {
    method: method(),
    methodLabel: methodLabel(),
    clientId: state.clientId,
    resource: "",
    authority: state.authority,
    redirectUri: state.redirectUri,
    account: state.accountUsername || username,
    naaAccountFound: Boolean(state.accountUsername),
    // `false` heisst: Die Bibliothek ist auf eine gewoehnliche Clientanwendung
    // zurueckgefallen. Die Seite soll messen, nicht die Erwartung wiederholen.
    naaBridge: state.naaBridge,
    employeeId,
    lastError: state.lastError,
  };
}

/* --------------------------------------------------------- Fehlerauswertung */

const INTERACTION_CODES = new Set([
  "interaction_required",
  "login_required",
  "consent_required",
  "no_account_error",
  "no_network_connectivity",
  "nested_app_auth_unavailable",
]);

function isInteractionRequired(error) {
  if (!error) return false;
  if (error.name === "InteractionRequiredAuthError") return true;
  return INTERACTION_CODES.has(String(error.errorCode || ""));
}

/**
 * Holt die Nachweisanforderung aus dem Fehler. Sie ist ein JSON-Text und wird
 * UNVERAENDERT weitergereicht - jede Umformung macht sie wertlos.
 */
function extractClaims(error) {
  if (error && typeof error.claims === "string" && error.claims.trim() !== "") {
    return error.claims;
  }
  return "";
}

function toAuthError(error) {
  state.lastError = describe(error);
  const code = String((error && error.errorCode) || "");
  if (code === "consent_required") return new AuthError("CLIENT_AUTH_CONSENT_REQUIRED", error);
  if (code === "popup_window_error" || code === "empty_window_error") {
    return new AuthError("CLIENT_AUTH_POPUP_BLOCKED", error);
  }
  if (code === "user_cancelled" || code === "invalid_grant" || isInteractionRequired(error)) {
    return new AuthError("CLIENT_AUTH_INTERACTION_REQUIRED", error);
  }
  return new AuthError("CLIENT_AUTH_FAILED", error);
}

/**
 * Kurzbeschreibung eines Fehlers fuer die Diagnoseseite.
 *
 * Bewusst knapp und ohne den Meldungstext der Bibliothek im Volltext: Ihre Meldungen
 * tragen gelegentlich Kontoname und Mandanten-Id, und diese Zeichenkette wird angezeigt.
 */
function describe(error) {
  if (!error) return "";
  const code = error.errorCode || error.code || error.name || "";
  return String(code).slice(0, 120);
}
