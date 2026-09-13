/**
 * Seite `#/anmeldung`: der Zugang zu TANSS.
 *
 * Sie ist der Regelweg, nicht die Ausnahme. Es gibt keinen Dienst mehr, der im Namen des
 * Technikers handeln koennte - er meldet sich selbst an und bekommt sein EIGENES Token.
 * Dessen Mitarbeiterbezug steht in der signierten Aussage und ist nicht faelschbar.
 *
 * Die Seite liegt bewusst NICHT hinter der Anmeldepruefung: Ein Anmeldefenster, das eine
 * Anmeldung voraussetzt, ist eine Sackgasse - der Benutzer kaeme aus ihr nur noch durch
 * Schliessen des Panes heraus.
 *
 * Die Zugangsdaten verlassen die Absendefunktion nicht und stehen ausschliesslich im
 * Rumpf der Anfrage, nie in der Adresse, wo sie in jedem Proxy-Protokoll landeten.
 */

import { T, errorText, t } from "../../i18n/de.js";
import * as api from "../api.js";
import * as auth from "../auth.js";
import { apiHost, config, setApiBase } from "../config.js";
import * as ui from "../ui.js";

/**
 * Meldet den Benutzer an TANSS an.
 *
 * Laeuft ueber denselben Verteiler wie alles andere und liefert denselben Umschlag: Es
 * fliegt nichts - Netzabbruch, zwischengeschaltete Anmeldeseite und falsche Zugangsdaten
 * kommen alle als `{ok:false}` zurueck.
 */
function tanssLogin(body, signal) {
  return api.post("api/auth/tanss", { json: body, signal });
}

/* ------------------------------------------------------------------------ Seite */

/**
 * @param {{root: HTMLElement, me: object|null, meError: object|null, signal: AbortSignal,
 *          navigate: Function, showError: Function, clearError: Function,
 *          reload: Function}} ctx
 * @returns {() => void} Aufraeumfunktion fuer den Routenwechsel.
 */
export function render(ctx) {
  const diagnostics = auth.diagnostics();
  const root = ui.el("div", { class: "stack" });

  if (ctx.me) {
    ui.mount(root, ...signedIn(ctx, diagnostics));
  } else {
    ui.mount(root, ...signedOut(ctx, diagnostics));
  }

  ui.replace(ctx.root, root);
  return () => {};
}

/** Bereits angemeldet: dann erklaert die Seite nur noch, WIE - und bietet das Abmelden an. */
function signedIn(ctx, diagnostics) {
  return [
    ui.banner({
      tone: "ok",
      label: ctx.me.name
        ? t("auth.signedInAs", { name: ctx.me.name })
        : t("auth.employeeId", { id: ctx.me.employeeId }),
    }),
    ui.kv([
      [T.diagnose.authMethod, diagnostics.methodLabel],
      [T.diagnose.authAccount, diagnostics.account],
    ]),
    // Abmelden betrifft ausschliesslich den TANSS-Rueckfall: das Microsoft-Konto gehoert
    // dem Outlook-Host, nicht uns - wir koennen und duerfen es nicht abmelden.
    auth.hasFallbackSession()
      ? ui.button({
          label: T.auth.signOut,
          variant: "danger",
          onClick: () => {
            auth.clearFallbackSession();
            void ctx.reload();
          },
        })
      : null,
    ui.button({ label: T.app.back, variant: "ghost", onClick: () => void ctx.reload() }),
  ];
}

/**
 * Nicht angemeldet.
 *
 * Die Erklaerung richtet sich nach dem, was tatsaechlich fehlgeschlagen ist - ein
 * allgemeiner Satz "Anmeldung fehlgeschlagen" laesst den Administrator raten, und geraten
 * wird hier an drei verschiedenen Stellen falsch.
 */
function signedOut(ctx, diagnostics) {
  const children = [];

  // Ohne TANSS-Adresse gibt es nichts anzumelden. Das kommt vor, wenn das Manifest sie
  // nicht mitbringt und keine config.json daneben liegt - dann fragt die Seite danach,
  // statt die Anmeldung gegen den Hoster laufen zu lassen und einen Netzfehler zu
  // zeigen, der nichts erklaert.
  if (!config().apiBase) {
    children.push(buildInstanceForm(ctx));
    return children;
  }

  // Gegen welche Instanz angemeldet wird, steht sichtbar da und wird nicht verschwiegen:
  // Der Techniker tippt gleich sein Kennwort ein, und bei einer gemeinsam genutzten
  // Ablage ist die Adresse das Einzige, was die Instanzen unterscheidet.
  children.push(ui.banner({ tone: "info", label: t("auth.targetHost", { host: apiHost() }) }));

  // Vorrang hat die Aussage des Dienstes: "kein Mitarbeiter zugeordnet" ist etwas voellig
  // anderes als "kein Token bekommen", und nur der Dienst kann das erste wissen.
  if (ctx.meError) {
    children.push(ui.banner({ tone: "warn", label: errorText(ctx.meError) }));
  }

  if (auth.hasEntraApp() && diagnostics.lastError) {
    children.push(ui.banner({ tone: "info", label: T.errors.CLIENT_AUTH_CONSENT_REQUIRED }));
  }

  children.push(
    ui.kv([
      [T.diagnose.authMethod, diagnostics.methodLabel],
      [T.diagnose.authClientId, diagnostics.clientId],
      [T.diagnose.authResource, diagnostics.resource],
      [T.diagnose.authAccount, diagnostics.account],
      [T.diagnose.authLastError, diagnostics.lastError],
    ]),
    ui.row([
      ui.button({ label: T.app.retry, variant: "primary", onClick: () => void ctx.reload() }),
      ui.button({
        label: T.nav.diagnose,
        variant: "ghost",
        onClick: () => ctx.navigate("#/diagnose"),
      }),
    ]),
    buildFallbackForm(ctx),
  );
  return children;
}

/**
 * Abfrage der TANSS-Adresse - die dritte und letzte Stufe.
 *
 * Angenommen wird nur eine https-Adresse, und nur eine aus der Positivliste, sofern die
 * Installation eine fuehrt. Das ist keine Formsache: Die naechste Maske fragt nach dem
 * Kennwort, und eine untergeschobene Adresse waere eine Anmeldemaske, die es woanders
 * hin schickt.
 */
function buildInstanceForm(ctx) {
  const box = ui.el("div", { class: "stack" });
  const adresse = ui.input({ placeholder: "https://tanss.example.de/backend" });

  function uebernehmen() {
    const wert = adresse.value.trim();
    if (wert === "") return;
    if (!setApiBase(wert)) {
      ctx.showError({ code: "CONFLICT", message: T.auth.instanceRejected });
      return;
    }
    void ctx.reload();
  }

  ui.mount(
    box,
    ui.section({
      heading: T.auth.instanceHeading,
      children: [
        ui.banner({ tone: "info", label: T.auth.instanceMissing }),
        ui.field({
          label: T.auth.instanceLabel,
          control: adresse,
          hint: T.auth.instanceHint,
          required: true,
        }),
        ui.button({ label: T.auth.instanceSave, variant: "primary", onClick: uebernehmen }),
      ],
    }),
  );

  adresse.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      uebernehmen();
    }
  });
  return box;
}


/**
 * Der TANSS-Anmelderueckfall.
 *
 * Er existiert ausschliesslich fuer eine Organisation ganz OHNE Entra und ist ab Werk
 * abgeschaltet; dann antwortet die Route mit HTTP 404, und das Formular ersetzt sich durch
 * den Satz, dass dieser Weg hier nicht vorgesehen ist. Die Sitzung, die dabei entsteht,
 * lebt im `sessionStorage` dieses Fensters - nie in den `roamingSettings`, denn die liegen
 * IM POSTFACH: wer es lesen darf, haette dort sonst ein Bearer-Token auf TANSS gefunden.
 */
function buildFallbackForm(ctx) {
  const box = ui.el("div", { class: "stack" });
  const userInput = ui.input({ autocomplete: "username" });
  const passwordInput = ui.input({ type: "password", autocomplete: "current-password" });
  const tokenInput = ui.input({ autocomplete: "one-time-code" });

  async function signIn() {
    const username = userInput.value.trim();
    if (username === "" || passwordInput.value === "") return;

    ctx.clearError();
    ui.setBusy(box, true);
    const result = await tanssLogin(
      { username, password: passwordInput.value, token: tokenInput.value.trim() },
      ctx.signal,
    );
    if (ctx.signal.aborted) return;
    ui.setBusy(box, false);
    // Das Kennwort verlaesst diese Funktion nicht und bleibt nicht im Formular stehen.
    passwordInput.value = "";

    if (result.ok) {
      // Die Sitzung liegt bereits: Sie entsteht im Anmeldevorgang selbst, samt
      // Erneuerungstoken und dem Zuschnitt des oertlichen Speichers auf diesen
      // Mitarbeiter. Hier ist nichts mehr zu uebernehmen.
      void ctx.reload();
      return;
    }
    ctx.showError(result.error);
  }

  ui.mount(
    box,
    ui.section({
      heading: T.auth.signInTanss,
      children: [
        ui.banner({ tone: "info", label: T.auth.fallbackNotice }),
        ui.field({ label: T.auth.username, control: userInput, required: true }),
        ui.field({ label: T.auth.password, control: passwordInput, required: true }),
        ui.field({
          label: T.auth.twoFactor,
          control: tokenInput,
          hint: T.auth.twoFactorHint,
        }),
        ui.button({ label: T.auth.signIn, variant: "primary", onClick: () => void signIn() }),
      ],
    }),
  );

  // Eingabe mit der Eingabetaste abschliessen - ohne `<form>`, weil ein Formular im
  // Taskpane die Seite neu laden und damit Office.js ein zweites Mal starten wuerde.
  for (const control of [userInput, passwordInput, tokenInput]) {
    control.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        void signIn();
      }
    });
  }
  return box;
}
