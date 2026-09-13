/**
 * Einstieg des Taskpanes: Office abwarten, Faehigkeiten pruefen, Route waehlen.
 *
 * Die wichtigste Eigenschaft dieser Datei ist nicht, was sie tut, sondern was sie verhindert:
 * einen Dauerladezustand. Ein `catch`, das selbst in einen `TypeError` laeuft - etwa weil
 * der Fehlerkoerper leer ist -, laesst den Ladekreis fuer immer stehen. Hier endet JEDER Pfad -
 * fehlendes office.js, zu alter Host, abgewiesene Anmeldung, kaputte Seitendatei - in einer
 * lesbaren deutschen Meldung mit einer Handlung daneben.
 *
 * Der Router ist bewusst schlicht: Hash-Routen, weil damit jede Manifest-URL dieselbe
 * `index.html` trifft und der Reverse Proxy keinen SPA-Rueckfall braucht.
 */

import { T, t } from "../i18n/de.js";
import * as api from "./api.js";
import * as auth from "./auth.js";
import * as office from "./office.js";
import * as ui from "./ui.js";

/** Mindestpaket. Darunter fehlt `userProfile.accountType`, und die Mailweiche traegt nicht. */
const REQUIRED_MAILBOX_SET = "1.6";

/**
 * Die fuenf Routen. `needs` ist die Art des geoeffneten Elements; `guarded` sagt, ob die
 * Seite eine aufgeloeste TANSS-Identitaet braucht - `#/diagnose` und `#/anmeldung`
 * ausdruecklich nicht, denn sie sind die Seiten fuer genau den Fall, dass sie fehlt.
 */
const ROUTES = new Map([
  ["#/ticket/neu", { module: "./pages/ticket-neu.js", title: T.nav.ticketNeu, needs: "message", guarded: true }],
  ["#/ticket/anhaengen", { module: "./pages/ticket-anhaengen.js", title: T.nav.ticketAnhaengen, needs: "message", guarded: true }],
  ["#/termin", { module: "./pages/termin.js", title: T.nav.termin, needs: "appointment", guarded: true }],
  ["#/anmeldung", { module: "./pages/anmeldung.js", title: T.nav.anmeldung, needs: null, guarded: false }],
  ["#/diagnose", { module: "./pages/diagnose.js", title: T.nav.diagnose, needs: null, guarded: false }],
]);

/** Zustand des Panes. Ein einziges Objekt - es gibt nichts, was zwei Besitzer haette. */
const app = {
  /** @type {HTMLElement|null} */ root: null,
  /** @type {HTMLElement|null} */ view: null,
  /** @type {HTMLElement|null} */ titleNode: null,
  /** @type {HTMLElement|null} */ subtitleNode: null,
  /** @type {object|null} */ errorBar: null,
  /** @type {object|null} */ me: null,
  /** @type {object|null} */ meError: null,
  /** @type {AbortController|null} */ pageAbort: null,
  /** @type {(() => void)|null} */ pageCleanup: null,
  itemType: null,
  route: "",
};

/* --------------------------------------------------------------------- Geruest */

/** Baut Kopf, Fehlerleiste und Inhaltsbereich. Danach steht das Geruest fuer immer. */
function buildShell() {
  const container = document.getElementById("app");
  if (!container) throw new Error("#app fehlt");
  ui.clear(container);

  app.titleNode = ui.el("h1", { class: "hdr-title", text: T.app.name });
  app.subtitleNode = ui.el("p", { class: "hdr-sub" });
  app.errorBar = ui.errorBar();
  app.view = ui.el("main", { class: "view" });

  const diagnoseButton = ui.button({
    label: T.nav.diagnose,
    variant: "ghost",
    onClick: () => navigate("#/diagnose"),
  });
  diagnoseButton.classList.add("btn--tiny");

  // Das Zeichen des Herausgebers, nicht das des Werkzeugs. Es steht im Kopf und damit
  // auch auf der Anmeldeseite - dort, wo der Techniker sein TANSS-Kennwort eintippt und
  // wissen soll, wessen Oberflaeche er vor sich hat. Im Menueband steht dagegen
  // weiterhin das "T": Ein Menuebandsymbol ist 16 Pixel gross, und darin zerfaellt eine
  // Bildmarke zu Sprenkeln, waehrend ein Buchstabe die FUNKTION benennt.
  const logo = ui.el("img", {
    class: "hdr-logo",
    attrs: { src: "assets/logo.png", alt: T.app.vendor, width: 72, height: 25 },
  });

  const header = ui.el("header", { class: "hdr" }, [
    logo,
    ui.el("div", { class: "hdr-main" }, [app.titleNode, app.subtitleNode]),
    diagnoseButton,
  ]);

  ui.mount(container, header, app.errorBar.root, app.view);
  app.root = container;
}

/** Kopfzeile: welche Seite, und als wer wir handeln. */
function setHeader(title, subtitle) {
  if (app.titleNode) app.titleNode.textContent = title;
  if (app.subtitleNode) app.subtitleNode.textContent = subtitle;
}

function identitySubtitle() {
  if (!app.me) return "";
  const name = app.me.name || "";
  return name ? t("auth.signedInAs", { name }) : t("auth.employeeId", { id: app.me.employeeId });
}

/**
 * Endzustand: eine Meldung, die erklaert, was nicht geht, und - wenn moeglich - eine
 * Handlung. Kein Ladekreis, kein leeres Pane.
 */
function fatal(message, actionLabel = "", onAction = null) {
  if (!app.view) {
    // Selbst das Geruest fehlt. Dann wenigstens ein Absatz, damit die Seite nicht weiss bleibt.
    const container = document.getElementById("app") || document.body;
    container.replaceChildren(ui.el("p", { class: "fatal", text: message }));
    return;
  }
  ui.replace(
    app.view,
    ui.el("div", { class: "fatal" }, [
      ui.el("p", { class: "fatal-text", text: message }),
      actionLabel && onAction ? ui.button({ label: actionLabel, onClick: onAction }) : null,
      ui.button({ label: T.nav.diagnose, variant: "ghost", onClick: () => navigate("#/diagnose") }),
    ]),
  );
}

/* ---------------------------------------------------------------------- Router */

/**
 * Aktuelle Route.
 *
 * Ohne Hash gilt die Standardroute nach Art des geoeffneten Elements. Ein unbekannter Hash
 * wird NICHT stillschweigend darauf umgebogen: er bedeutet, dass Manifest und Taskpane
 * auseinanderlaufen, und das gehoert gesagt statt kaschiert.
 */
function currentRoute() {
  // Nachgestellte Parameter (`#/ticket/neu?x=1`) werden abgeschnitten - Zustand gehoert
  // nicht in die Adresse, sondern in den Vorgang.
  const clean = (window.location.hash || "").split("?")[0];
  if (clean === "" || clean === "#" || clean === "#/") return defaultRoute();
  return clean;
}

/** Standardroute nach Art des geoeffneten Elements. */
function defaultRoute() {
  if (app.itemType === "appointment") return "#/termin";
  if (app.itemType === "message") return "#/ticket/neu";
  return "#/diagnose";
}

export function navigate(route) {
  if (window.location.hash === route) {
    void renderRoute();
    return;
  }
  window.location.hash = route;
}

/**
 * Zeigt die Seite zur aktuellen Route.
 *
 * Vor jedem Wechsel wird die alte Seite abgebrochen: ihr `AbortSignal` faellt, damit eine
 * noch laufende Suche nicht in eine abgeraeumte Oberflaeche schreibt, und ihre
 * Aufraeumfunktion laeuft. Ohne das haengen entprellte Zeitgeber und Ereignisbindungen
 * an Knoten, die es nicht mehr gibt.
 */
async function renderRoute() {
  const route = currentRoute();
  app.route = route;

  if (app.pageAbort) app.pageAbort.abort();
  if (app.pageCleanup) {
    try {
      app.pageCleanup();
    } catch {
      // Eine fehlerhafte Aufraeumfunktion darf den Wechsel nicht verhindern.
    }
    app.pageCleanup = null;
  }
  app.errorBar.hide();

  const entry = ROUTES.get(route);
  if (!entry) {
    setHeader(T.app.name, "");
    fatal(`${T.nav.unknownRoute} ${T.nav.unknownRouteHint}`);
    return;
  }

  setHeader(entry.title, identitySubtitle());

  if (entry.needs && entry.needs !== app.itemType) {
    fatal(
      entry.needs === "message" ? T.host.wrongItemMessage : T.host.wrongItemAppointment,
      T.nav.diagnose,
      () => navigate("#/diagnose"),
    );
    return;
  }

  if (entry.guarded && !app.me) {
    // Ohne Anmeldung gibt es keine fachliche Seite - aber sehr wohl einen Weg dorthin.
    // Es gibt genau EINEN: die Anmeldung an TANSS. Frueher stand hier zusaetzlich die
    // Frage, ob eine Entra-Anwendung hinterlegt ist; die entscheidet inzwischen nur
    // noch darueber, ob die Nachricht im Originalformat geholt werden kann, und hat
    // mit dem Zugang nichts mehr zu tun.
    const error = app.meError || { code: "UNAUTHENTICATED" };
    if (error.code === "UNAUTHENTICATED") {
      navigate("#/anmeldung");
      return;
    }
    app.errorBar.show(error);
    fatal(error.message || T.errors.UNAUTHENTICATED, T.app.retry, () => void reload());
    return;
  }

  app.pageAbort = new AbortController();
  ui.replace(app.view, ui.loading(T.app.loading));

  try {
    const page = await import(entry.module);
    if (app.route !== route) return; // In der Zwischenzeit wurde weitergeklickt.
    ui.clear(app.view);
    const cleanup = await page.render(pageContext(route));
    if (typeof cleanup === "function") app.pageCleanup = cleanup;
  } catch (error) {
    // Auch ein Fehler IN einer Seite endet sichtbar - und geht als eine Zeile ins Journal.
    showError({ code: "CLIENT_UNEXPECTED" });
    fatal(T.errors.CLIENT_UNEXPECTED, T.app.retry, () => void renderRoute());
    void api.reportClientError(route, describeError(error));
  }
}

/** Der Vertrag, den jede Seite unter `js/pages/` bekommt. */
function pageContext(route) {
  return {
    root: app.view,
    me: app.me,
    meError: app.meError,
    route,
    signal: app.pageAbort.signal,
    navigate,
    showError,
    clearError: () => app.errorBar.hide(),
    setSubtitle: (value) => setHeader(ROUTES.get(route).title, value),
    reload,
  };
}

export function showError(error) {
  app.errorBar.show(error);
}

/* ------------------------------------------------------------------ Anmeldung */

/**
 * Holt `GET /api/me` - Identitaet, Grenzen, Vorgaben, Funktionen.
 *
 * Ein Aufruf beim Oeffnen des Panes, und danach kennt die Oberflaeche alles, was sie
 * braucht. Jeder zusaetzliche Aufruf kostete eine Umlaufzeit, bevor der Techniker das
 * erste Feld sieht.
 */
async function loadMe() {
  const result = await api.get("api/me");
  if (result.ok) {
    app.me = result.data;
    app.meError = null;
  } else {
    app.me = null;
    app.meError = result.error;
  }
}

/** Identitaet neu holen und die aktuelle Seite neu zeichnen. */
export async function reload() {
  await loadMe();
  await renderRoute();
}

/* --------------------------------------------------------------------- Start */

/** Kurzbeschreibung eines Fehlers fuer das Journal - nie ein Mailinhalt, nie ein Token. */
function describeError(error) {
  if (!error) return "";
  const name = error.name || "Error";
  const code = error.code ? `:${error.code}` : "";
  return `${name}${code}`.slice(0, 200);
}

/**
 * Faehigkeitspruefung.
 *
 * Sie fragt den laufenden Host, statt aus einer Build-Nummer zu schliessen. Faellt sie
 * negativ aus, ist das kein Fehler des Benutzers, und die Meldung sagt genau, was fehlt
 * und was daran haengt.
 */
function checkCapabilities() {
  if (!office.hasMailbox()) {
    fatal(T.host.notOutlook);
    return false;
  }
  if (!office.isSetSupported("Mailbox", REQUIRED_MAILBOX_SET)) {
    fatal(t("host.setTooOld", { required: REQUIRED_MAILBOX_SET }));
    return false;
  }
  return true;
}

async function start() {
  buildShell();
  setHeader(T.app.name, T.boot.officeWaiting);

  try {
    await office.onReady();
  } catch {
    fatal(T.host.notOutlook);
    return;
  }

  setHeader(T.app.name, T.boot.hostCheck);
  if (!checkCapabilities()) return;

  app.itemType = office.getItemType();

  // MSAL vorbereiten, bevor die erste Anfrage sie braucht. Der Aufruf wirft nicht; ob eine
  // Entra-Anwendung hinterlegt ist, entscheidet danach `auth.hasEntraApp()`.
  setHeader(T.app.name, T.boot.signingIn);
  await auth.initialize();

  setHeader(T.app.name, T.boot.profileLoading);
  await loadMe();

  window.addEventListener("hashchange", () => {
    void renderRoute();
  });
  await renderRoute();
}

/**
 * Auffangnetz fuer alles, was an den Seiten vorbeigeht.
 *
 * Eine Zeile ins Journal des eigenen Dienstes - Route, Fehlername, Hostname. Mehr nicht:
 * kein Betreff, kein Mailtext, kein Token, und kein anderer Empfaenger als dieser Dienst.
 */
window.addEventListener("error", (event) => {
  void api.reportClientError(app.route || window.location.hash, describeError(event.error));
});

window.addEventListener("unhandledrejection", (event) => {
  void api.reportClientError(app.route || window.location.hash, describeError(event.reason));
});

start().catch((error) => {
  fatal(T.errors.CLIENT_UNEXPECTED);
  void api.reportClientError("start", describeError(error));
});
