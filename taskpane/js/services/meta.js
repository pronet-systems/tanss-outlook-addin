/**
 * Identitaet, Vorbelegungen, Anmeldung - was frueher der Dienst beantwortet hat.
 *
 * `me()` ist das Herzstueck: Es sagt beim Oeffnen des Panes in EINEM Vorgang, wer der
 * Benutzer in TANSS ist, welche Grenzen gelten und was vorbelegt wird. Ein Aufruf statt
 * mehrerer ist der Grund, warum die Oberflaeche ohne Rahmenwerk auskommt - es gibt
 * keinen verteilten Zustand zu binden, wenn alles Abhaengige zusammen kommt.
 *
 * Die Identitaet wird dabei NICHT aufgeloest und schon gar nicht geraten: Sie steht
 * durch die Anmeldung fest. Das ist der grosse Unterschied zum Dienstbetrieb, in dem
 * ein Postfach erst einem Mitarbeiter zugeordnet werden musste - mit allen
 * Mehrdeutigkeiten, die daran hingen.
 */

import { reasonOf } from "../tanss/errors.js";
import { VERSION } from "../version.js";
import { rebuildStore, repository, session, settings, store } from "../runtime.js";

/**
 * Wer bin ich, was gilt, was ist vorbelegt.
 *
 * Jeder Teil ist entbehrlich: Faellt der Anzeigename aus, steht ein Satz in `warnings`,
 * und das Pane oeffnet trotzdem. Was hier scheitern darf, darf das Oeffnen nicht
 * verhindern.
 */
export async function me() {
  const config = settings();
  const current = session();
  const warnings = [];

  return {
    employeeId: current.employeeId(),
    name: await displayName(current, warnings),
    version: VERSION,
    instance: config.instance,
    tanssFrontendUrl: config.frontendBase,
    limits: {
      maxEmlBytes: config.limits.maxEmlBytes,
    },
    features: {
      appointments: config.features.appointments,
      createSupport: config.features.createSupport,
      // Ohne hinterlegte Anwendungs-Id gibt es keinen Weg zur Originalnachricht - der
      // Schalter sagt das ehrlich, statt einen Weg anzubieten, der nicht existiert.
      graphMail: config.features.graphMail && Boolean(config.entra.clientId),
      graphCalendar: config.features.graphCalendar && Boolean(config.entra.clientId),
    },
    defaults: {
      supportTypeId: config.defaults.supportTypeId || null,
      location: config.defaults.location,
      internal: config.defaults.internal,
      ticketTypeId: store().intPref("ticketTypeId") || config.defaults.ticketTypeId || null,
    },
    prefs: store().prefs(),
    warnings,
  };
}

/**
 * Der Anzeigename aus dem Verzeichnis - oder der Anmeldename.
 *
 * Der Name ist reine Anzeige und wird bewusst nirgends aufbewahrt: Er ist aus TANSS
 * jederzeit lesbar, und eine gespeicherte Kopie waere ein zweiter, alternder
 * Wahrheitsanspruch. Faellt das Verzeichnis aus, zeigt das Pane den Anmeldenamen -
 * das ist immer noch besser als eine nackte Nummer.
 */
async function displayName(current, warnings) {
  const id = current.employeeId();
  try {
    for (const employee of await repository().listTechnicians()) {
      if (employee.id === id) return employee.name;
    }
  } catch (error) {
    // Der Grund gehoert in die Meldung. "War nicht abrufbar" allein hat diesen Ausfall
    // schon zweimal ueberdauert, ohne dass jemand sagen konnte, woran es liegt.
    warnings.push(
      "Die Technikerliste war nicht abrufbar. Der Name wird deshalb nicht angezeigt; "
      + "auf die Ticketanlage hat das keinen Einfluss. "
      + `Grund: /api/tanss.x/v1/technicians — ${reasonOf(error)}`,
    );
  }
  return current.username();
}

/**
 * Merkt eine Vorbelegung.
 *
 * Unbekannte Schluessel werden verworfen, nicht angelegt - der oertliche Speicher ist
 * beliebig beschreibbar, und ein offener Schluesselraum waere ein kleiner, kostenloser
 * Datenspeicher fuer alles, was jemand spaeter darin ablegen moechte.
 */
export async function setPref({ key, value }) {
  const accepted = store().setPref(key, value);
  return { key, accepted };
}

/**
 * Anmeldung an TANSS.
 *
 * Danach wird der oertliche Speicher neu zugeschnitten: Sein Bereich haengt an der
 * Mitarbeiternummer, und die steht erst jetzt fest. Ohne diesen Schritt sieht der
 * gerade Angemeldete die Vorbelegungen des zuvor Angemeldeten.
 */
export async function login({ username, password, token = "" }) {
  const result = await session().login({ username, password, token });
  rebuildStore();
  return {
    employeeId: result.employeeId,
    displayName: result.username,
  };
}

/** Meldet ab und raeumt die Sitzung weg. */
export async function logout() {
  session().clear();
  rebuildStore();
  return null;
}

/**
 * Ein unerwarteter Clientfehler.
 *
 * Im Dienstbetrieb ging hier eine Zeile ins Journal. Es gibt kein Journal mehr, und es
 * wird auch keines erfunden: Ein Fehlerbericht an einen Dritten waere ein vierter
 * Empfaenger fuer Daten, die durch dieses Pane laufen - Betreff, Klartext, Token. Der
 * Fehler bleibt deshalb auf diesem Rechner, in der Konsole des Browsers, wo ihn der
 * abholen kann, der ihn sucht.
 */
export async function reportClientError({ route, message }) {
  // eslint-disable-next-line no-console
  console.warn(`[tanss-addin] ${route}: ${message}`);
  return null;
}
