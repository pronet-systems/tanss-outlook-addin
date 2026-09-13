/**
 * Seite `#/diagnose`: was der laufende Host tatsaechlich meldet.
 *
 * Der Entwurf dieses Add-ins steht an mehreren Stellen auf Dokumentationswissen - welches
 * Anforderungspaket welche Funktion traegt, ob NAA auf diesem Client greift, ob das
 * Postfach in Exchange Online oder auf einem lokalen Exchange liegt. Dokumentationswissen
 * ist Erwartung; diese Seite ist Messung. Sie ist deshalb die erste Anlaufstelle, wenn
 * etwas nicht funktioniert, und sie ist bewusst OHNE Identitaetspruefung erreichbar: die
 * haeufigste Frage an sie lautet, warum die Anmeldung gerade nicht geht.
 *
 * Alles steht als Schluessel-Wert-Liste untereinander, damit ein Bildschirmfoto aus dem
 * Pane fuer den Administrator brauchbar ist. Der Knopf "Als Text kopieren" liefert
 * denselben Inhalt in die Zwischenablage - aus DERSELBEN Quelle, damit Bild und Text nie
 * auseinanderlaufen.
 *
 * Was hier ausdruecklich NICHT steht: kein Token, kein Tokenauszug, keine Mailinhalte,
 * keine Betreffzeile. `auth.diagnostics()` liefert bereits nur Kennungen, und die
 * Fehlerbeschreibung von MSAL ist dort auf den Fehlercode beschnitten - MSAL-Meldungen
 * tragen gelegentlich Kontoname und Mandanten-Id im Volltext.
 */

import { T, errorText, formatBytes } from "../../i18n/de.js";
import * as auth from "../auth.js";
import * as boot from "../boot.js";
import { apiHost, config, note as configNote } from "../config.js";
import * as mime from "../mime.js";
import * as office from "../office.js";
import * as ui from "../ui.js";

/** Trennzeichen innerhalb einer Zeile. Satzzeichen, kein uebersetzbarer Text. */
const SEPARATOR = " · ";

/** Anzeigewert. `undefined`, `null` und Leerstring werden zu einem ehrlichen Platzhalter. */
function show(value) {
  if (value === true) return T.app.yes;
  if (value === false) return T.app.no;
  if (value === null || value === undefined || value === "") return T.diagnose.notAvailable;
  return String(value);
}

/* ------------------------------------------------------------------------ Seite */

/**
 * @param {{root: HTMLElement, me: object|null, meError: object|null, signal: AbortSignal,
 *          navigate: Function, showError: Function, clearError: Function,
 *          reload: Function}} ctx
 * @returns {Promise<() => void>} Aufraeumfunktion fuer den Routenwechsel.
 */
export async function render(ctx) {
  ui.replace(ctx.root, ui.loading(T.app.loading));

  // Keine Probe gegen eine Gegenstelle. Es gibt keinen eigenen Dienst mehr, den man
  // fragen koennte, und ein Aufruf an TANSS setzte eine Anmeldung voraus - die gerade
  // fehlt, wenn diese Seite geoeffnet wird. Gezeigt wird stattdessen, was oertlich gilt.
  const settings = config();

  const host = office.hostInfo();
  const profile = office.getUserProfile();
  const diagnostics = auth.diagnostics();
  const me = ctx.me || null;
  const limits = me && me.limits ? me.limits : {};
  const features = me && me.features ? me.features : {};

  /* ------------------------------------------------------------------ Abschnitte */

  /**
   * Der Start - der erste Abschnitt, weil er die erste Frage beantwortet.
   *
   * Bleibt das Pane im Ladezustand, entscheiden diese vier Zeilen zwischen drei ganz
   * verschiedenen Ursachen: Die Bibliothek wurde gar nicht geladen (dann steht "nein" in
   * Zeile eins und meist ein Eintrag in Zeile vier), Outlook hat sich nicht gemeldet (dann
   * "ja" oben und "nein" darunter), oder die Kennung des Hosts fehlt in der Adresse - dann
   * weiss office.js selbst nicht, welchen Wirt es vor sich hat.
   */
  const bootBlocked = boot.blocked();
  const bootPairs = [
    [T.diagnose.bootOffice, show(boot.hadOfficeLibrary())],
    ...boot.entries(),
    [T.diagnose.bootHostInfo, boot.hostInfoParam() || T.diagnose.bootHostInfoMissing],
    [
      T.diagnose.bootCsp,
      bootBlocked.length === 0 ? T.diagnose.bootCspNone : bootBlocked.join(SEPARATOR),
    ],
  ];

  const hostPairs = [
    [T.diagnose.host, show(host.hostName || host.host)],
    [T.diagnose.platform, show(host.platform)],
    [T.diagnose.officeVersion, show(host.hostVersion)],
    [T.diagnose.requirementSets, show(office.supportedMailboxSets().join(SEPARATOR))],
    [
      T.diagnose.accountType,
      profile && profile.accountType
        ? `${profile.accountType}${SEPARATOR}${
            office.isEnterpriseMailbox()
              ? T.diagnose.accountTypeEnterprise
              : T.diagnose.accountTypeCloud
          }`
        : T.diagnose.notAvailable,
    ],
    [T.diagnose.mailbox, show(profile ? profile.emailAddress : "")],
    [T.diagnose.featureGraphMail, show(mime.currentSource() === "graph")],
  ];

  const authPairs = [
    [T.diagnose.authMethod, show(diagnostics.methodLabel)],
    [T.diagnose.authClientId, show(diagnostics.clientId)],
    [T.diagnose.authResource, show(diagnostics.resource)],
    [T.diagnose.authAccount, show(diagnostics.account)],
    [T.diagnose.authLastError, show(diagnostics.lastError)],
  ];

  /**
   * Was oertlich gilt.
   *
   * Diese vier Angaben beantworten die Fragen, die bei einer Stoerung tatsaechlich gestellt
   * werden: gegen welche Instanz das Pane arbeitet, unter welchem Kennzeichen es seine
   * Anmeldung ablegt, mit welcher Anwendungs-Id es sich bei Microsoft meldet - und ob die
   * Konfigurationsdatei ueberhaupt gelesen wurde. Ein Tippfehler im Manifest steht hier im
   * Klartext, statt sich als "nicht erreichbar" zu tarnen.
   */
  const configPairs = [
    [T.diagnose.configApi, show(apiHost())],
    [T.diagnose.configInstance, show(settings.instance)],
    [T.diagnose.configEntra, show(settings.entra ? settings.entra.clientId : "")],
    [T.diagnose.configSource, show(configNote())],
  ];

  const servicePairs = [
    [T.diagnose.serviceVersion, show(me ? me.version : "")],
    [
      T.diagnose.employee,
      me ? `${me.employeeId}${me.name ? SEPARATOR + me.name : ""}` : T.diagnose.notAvailable,
    ],
    [
      T.diagnose.maxEml,
      limits.maxEmlBytes ? formatBytes(limits.maxEmlBytes) : T.diagnose.notAvailable,
    ],
    [T.diagnose.featureAppointments, show(features.appointments === true)],
    [T.diagnose.featureCreateSupport, show(features.createSupport === true)],
  ];

  /* -------------------------------------------------------------------- Hinweise */

  /**
   * Der zuletzt aufgetretene Fehler - der Grund, aus dem diese Seite meist geoeffnet wird.
   *
   * Angezeigt wird der fertige deutsche Satz aus `i18n`, nicht der Rohtext.
   */
  const problems = [];
  if (ctx.meError) problems.push(errorText(ctx.meError));
  const warnings = me && Array.isArray(me.warnings) ? me.warnings : [];

  /* ------------------------------------------------------------------- Kopieren */

  const copyState = ui.el("p", { class: "muted" });

  /** Derselbe Inhalt als Text - aus denselben Paaren, damit nichts auseinanderlaeuft. */
  function asPlainText() {
    const lines = [];
    for (const group of [bootPairs, hostPairs, authPairs, configPairs, servicePairs]) {
      for (const [key, value] of group) lines.push(`${key}: ${value}`);
      lines.push("");
    }
    for (const problem of problems.concat(warnings)) lines.push(problem);
    return lines.join("\n");
  }

  async function copyAll() {
    try {
      await navigator.clipboard.writeText(asPlainText());
      copyState.textContent = T.diagnose.copied;
    } catch {
      // Manche Hosts sperren die Zwischenablage im Taskpane. Dann bleibt das
      // Bildschirmfoto - deshalb steht ohnehin alles sichtbar auf der Seite.
      copyState.textContent = T.diagnose.copyFailed;
    }
  }

  /* --------------------------------------------------------------------- Aufbau */

  ui.replace(
    ctx.root,
    ui.el("div", { class: "stack" }, [
      ui.el("p", { class: "muted", text: T.diagnose.intro }),

      ...problems.map((problem) => ui.banner({ tone: "warn", label: problem })),
      ...warnings.map((warning) => ui.banner({ tone: "warn", label: warning })),

      ui.section({ heading: T.diagnose.bootSection, children: [ui.kv(bootPairs)] }),
      ui.section({ heading: T.diagnose.host, children: [ui.kv(hostPairs)] }),
      ui.section({ heading: T.nav.anmeldung, children: [ui.kv(authPairs)] }),
      ui.section({ heading: T.diagnose.configuration, children: [ui.kv(configPairs)] }),
      ui.section({ heading: T.diagnose.service, children: [ui.kv(servicePairs)] }),

      ui.row([
        ui.button({ label: T.diagnose.copy, variant: "ghost", onClick: () => void copyAll() }),
        ui.button({ label: T.app.retry, variant: "ghost", onClick: () => void ctx.reload() }),
      ]),
      copyState,
    ]),
  );

  return () => {};
}
