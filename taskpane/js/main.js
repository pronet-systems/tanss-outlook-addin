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

// ZUERST: Das Startprotokoll haengt seinen Zuhoerer fuer Verstoesse gegen die
// Inhaltsrichtlinie an, und es zaehlt jede Millisekunde, die vorher vergeht. Module werden
// in der Reihenfolge ihrer Import-Zeilen ausgewertet - diese Zeile gehoert deshalb nach
// oben und nicht in die alphabetische Ordnung darunter.
import * as boot from "./boot.js";

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
  /**
   * Die Seite, von der aus die aktuelle betreten wurde - der Rueckweg der Nebenseiten.
   *
   * `#/diagnose` und `#/anmeldung` sind von ueberall her erreichbar. Ihr "Zurueck" fuehrte
   * bisher auf die Standardroute nach Art des geoeffneten Elements, und das ist bei einer
   * Mail immer "Ticket erstellen": Wer die Diagnose aus "An Ticket anhaengen" geoeffnet
   * hatte, landete beim Zurueckgehen in der falschen Maske - mitsamt dem Verlust dessen,
   * was er dort schon ausgewaehlt hatte.
   */
  previousRoute: "",
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
 *
 * `facts` traegt Messwerte, die zur Meldung gehoeren. Sie stehen hier und nicht nur auf
 * der Diagnoseseite, weil die Seitenmodule ueber einen dynamischen Import kommen und
 * deshalb aus dem Zwischenspeicher stammen koennen, waehrend diese Datei frisch geladen
 * wurde. Wer eine Stoerung untersucht, soll nicht davon abhaengen, welches Modul gerade
 * aktuell ist.
 *
 * @param {Array<[string, string]>} facts
 */
function fatal(message, actionLabel = "", onAction = null, facts = []) {
  // Ein Endzustand darf keinen Fortschrittstext stehen lassen. Sonst steht ueber der
  // Fehlermeldung weiter "Outlook wird abgewartet …" und behauptet, es laufe noch etwas -
  // genau die Verwechslung, die den ersten Lauf in echtem Outlook unlesbar gemacht hat.
  if (app.subtitleNode && Object.values(T.boot).includes(app.subtitleNode.textContent)) {
    app.subtitleNode.textContent = "";
  }
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
      facts.length > 0 ? ui.kv(facts) : null,
      actionLabel && onAction ? ui.button({ label: actionLabel, onClick: onAction }) : null,
      ui.button({ label: T.nav.diagnose, variant: "ghost", onClick: () => navigate("#/diagnose") }),
    ]),
  );
}

/**
 * Die Zeile ueber abgewiesene Ladevorgaenge.
 *
 * "nichts" darf nur dastehen, wenn auch wirklich zugehoert wurde. Fiel der fruehe Zuhoerer
 * aus, ist die leere Liste keine Entwarnung, sondern eine Luecke - und die wird benannt.
 */
function cspZeile(liste) {
  if (liste.length > 0) return liste.join(" \u00b7 ");
  return boot.hadEarlyListener() ? T.diagnose.bootCspNone : T.diagnose.bootCspLate;
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

  // Nur ein WECHSEL zaehlt als Herkunft. Ein erneutes Zeichnen derselben Seite - der Knopf
  // im Kopf auf der Diagnose, ein `reload()` nach der Anmeldung - darf den Rueckweg nicht
  // auf sich selbst umbiegen; sonst fuehrt "Zurueck" auf die Seite, auf der man steht.
  if (app.route && app.route !== route) app.previousRoute = app.route;
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

    // Zwei Zustaende, ein Ziel: Es liegt keine brauchbare Anmeldung vor.
    //
    //   - `UNAUTHENTICATED` - TANSS hat das Token abgewiesen, oder es lag keines vor.
    //   - Eine gescheiterte ERNEUERUNG (`detail: "renew"`, siehe `tanss/session.js`).
    //     Sie heisst dasselbe: Das Zugriffstoken ist abgelaufen und liess sich nicht
    //     ersetzen. Ob die Erneuerung am Netz scheiterte oder daran, dass die Instanz
    //     den Kopf `refreshToken` nicht durch ihre Vorabfrage laesst, ist von hier aus
    //     nicht zu erkennen - der Browser meldet beides als gewoehnlichen Netzfehler.
    //     Fuer den Techniker macht es auch keinen Unterschied: Er kommt in beiden
    //     Faellen nur ueber eine Anmeldung weiter.
    //
    // Frueher endete der zweite Fall in einer Fehlermeldung mit dem Knopf "Erneut
    // versuchen". Der war eine Sackgasse: Laesst die Instanz die Erneuerung nicht zu,
    // scheitert jeder weitere Versuch aus demselben Grund, und aus diesem Zustand kam
    // der Techniker nur noch durch Loeschen der Seitendaten heraus.
    //
    // Die Anmeldeseite VERWIRFT die Sitzung nicht. Das ist der Unterschied zum
    // Abmelden, und er ist beabsichtigt: War TANSS nur kurz nicht erreichbar, genuegt
    // dort "Erneut versuchen" - und ein kurzer Netzaussetzer kostet kein Kennwort.
    if (error.code === "UNAUTHENTICATED" || error.detail === "renew") {
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
    /**
     * Der Rueckweg einer Nebenseite: dorthin, woher sie betreten wurde.
     *
     * Die Seiten bestimmen ihn nicht selbst - sie wissen ihn gar nicht. Ohne Herkunft
     * bleibt die Standardroute nach Art des geoeffneten Elements; das ist der Fall, in
     * dem das Pane direkt auf einer Nebenseite geoeffnet wurde.
     */
    back: () => navigate(app.previousRoute || "#/"),
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
  // Der Router wird VOR dem Warten auf Outlook verdrahtet. Meldet sich der Host nicht,
  // haengt der Start genau hier - und genau dann wird die Diagnoseseite gebraucht. Stuende
  // die Verdrahtung wie frueher erst darunter, waere sie im Fehlerfall nicht erreichbar,
  // und der Knopf im Kopf taete nichts.
  window.addEventListener("hashchange", () => {
    void renderRoute();
  });

  setHeader(T.app.name, T.boot.officeWaiting);

  // Die Zeile wird gesetzt, BEVOR gewartet wird. Wer die Diagnoseseite waehrend des
  // Wartens oeffnet - der haeufigste Fall -, soll dort nicht eine fehlende Zeile
  // vorfinden, sondern den Satz "es laeuft noch". Eine Luecke liest sich wie ein Versehen.
  boot.mark(T.diagnose.bootReady, T.boot.officeWaiting);

  try {
    const info = await office.onReady();
    boot.mark(
      T.diagnose.bootReady,
      [T.app.yes, info.host || "?", info.platform || "?", `${boot.sinceStart()} ms`].join(" · "),
    );
  } catch (error) {
    // Der Grund wird festgehalten, BEVOR die Meldung erscheint: `detail` unterscheidet
    // "office.js fehlt" von "ohne Antwort", und genau diese Unterscheidung sucht hinterher
    // derjenige, der das Bildschirmfoto der Diagnoseseite bekommt.
    boot.mark(
      T.diagnose.bootReady,
      [T.app.no, (error && error.detail) || "?", `${boot.sinceStart()} ms`].join(" · "),
    );
    // Zwei Faelle, zwei Saetze. Fehlt office.js, laeuft die Seite ausserhalb von Outlook.
    // Ist die Bibliothek da und antwortet trotzdem nicht, ist die Einbettung gestoert -
    // eine ganz andere Suche.
    //
    // Darunter stehen die drei Messwerte, die diese Suche entscheiden: Fehlt die Kennung
    // des Hosts in der Adresse, weiss office.js selbst nicht, welchen Wirt es vor sich hat
    // und kann sich gar nicht melden. Steht dagegen etwas in der dritten Zeile, hat die
    // Inhaltsrichtlinie einen Ladevorgang abgewiesen, den office.js braucht. Beides fuehrt
    // zum selben Stillstand und hat nichts miteinander zu tun.
    const blockiert = boot.blocked();
    fatal(
      office.hasOfficeApi() ? T.host.officeSilent : T.host.notOutlook,
      T.app.retry,
      () => globalThis.location.reload(),
      [
        [T.diagnose.bootOffice, office.hasOfficeApi() ? T.app.yes : T.app.no],
        [T.diagnose.bootHostInfo, boot.hostInfoLine()],
        [T.diagnose.bootCsp, cspZeile(blockiert)],
      ],
    );
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
