/**
 * Der Outlook-Termin und sein TANSS-Einsatz.
 *
 * Dieses Modul reichert einen Termin um TANSS-Felder an. Es verschiebt ihn nicht, es
 * macht aus einem Terminvorschlag keinen festen Termin, und es fasst die Teilnehmer
 * nicht an. Was es nicht schreibt, kann es auch nicht loeschen - und genau daran haengt
 * die Vertraeglichkeit mit einer Kalendersynchronisation, die gegen dieselbe Instanz
 * laeuft.
 *
 * Die Kopplung selbst liegt an DREI Stellen, in dieser Rangfolge:
 *
 * 1. am Termin, als Eigenschaft des Kalendereintrags - sie wandert mit dem Postfach,
 * 2. in TANSS, als Kopplungswert am Einsatz,
 * 3. im oertlichen Speicher dieses Rechners, als blosser Zwischenspeicher.
 *
 * Die dritte darf jederzeit verloren gehen; das kostet eine Suchanfrage, keine
 * Zuordnung.
 */

import { ApiError } from "../tanss/errors.js";
import { repository, session, settings, store } from "../runtime.js";

/** Die drei Einsatzorte sind eine feste Aufzaehlung, keine Liste aus der Instanz. */
const LOCATIONS = ["OFFICE", "ONSITE", "REMOTE"];

/** Wie viele offene Tickets der Firma die Terminmaske zur Auswahl anbietet. */
const TICKET_LIMIT = 25;

/**
 * Den Einsatz zum geoeffneten Termin finden.
 *
 * Ohne Elementkennung gibt es nichts, woran eine Kopplung festzumachen waere: Ein noch
 * nie gespeicherter Termin hat keine. Die Oberflaeche bittet dann zuerst um "Speichern",
 * statt eine Zuordnung zu raten.
 */
export async function resolve(query, { signal } = {}) {
  const itemKey = String(query.itemKey || "").trim();
  if (!itemKey) throw new ApiError("APPOINTMENT_NOT_SAVED");

  const start = query.start ? new Date(query.start) : null;
  if (start && Number.isNaN(start.getTime())) {
    throw new ApiError("CONFLICT", "Der Beginn des Termins war nicht lesbar.");
  }

  // Stufe 1: die am Termin hinterlegte Kennung, sonst der Zwischenspeicher. Welche der
  // beiden es war, gehoert in die Antwort - im Betrieb erklaert der Unterschied, warum
  // eine Kopplung nach einem Rechnerwechsel ploetzlich fehlt.
  let supportId = Number(query.supportId) || 0;
  let fromCache = false;
  if (!supportId) {
    const link = store().appointmentLink(itemKey);
    if (link) {
      supportId = link.supportId;
      fromCache = true;
    }
  }

  const match = await repository().resolveSupport({
    employeeId: session().employeeId(),
    supportId: supportId || null,
    uid: query.icalUid || "",
    start,
    durationMinutes: Number(query.durationMinutes) || 0,
  }, { signal });

  const source = match.source === "property" && fromCache ? "link" : match.source;

  if (match.supportId) {
    store().rememberAppointmentLink(itemKey, {
      supportId: match.supportId,
      uid: query.icalUid || "",
      startUnix: start ? Math.floor(start.getTime() / 1000) : 0,
      source,
    });
  }

  return {
    support: supportPayload(match),
    candidates: match.candidates,
    source,
    readOnly: match.readOnly,
    // Der Altmarker im Terminkoerper wird nicht mehr gelesen: Er steht im sichtbaren
    // Text und ist damit von jedem Teilnehmer beschreibbar. Die Kopplung kommt aus den
    // drei Stufen oben, nie von dort.
    legacyMarker: { found: false, prefill: {} },
    canCreate: !match.supportId && settings().features.createSupport,
    metaInconclusive: match.metaInconclusive,
  };
}

/**
 * Der gefundene Einsatz, so weit die Terminmaske ihn braucht.
 *
 * `text` reist mit, und das ist kein Beiwerk: Die Maske zeigt den vorhandenen Text zum
 * Bearbeiten an - ohne ihn ueberschriebe der erste Speichervorgang eine bestehende
 * Beschreibung mit Leere.
 */
function supportPayload(match) {
  if (!match.supportId || !match.support) return null;
  const support = match.support;
  const start = support.date > 0 ? new Date(support.date * 1000) : null;
  return {
    supportId: support.id,
    companyId: support.companyId,
    companyName: match.companyName || "",
    ticketId: support.ticketId,
    typeId: support.typeId,
    location: support.location,
    internal: support.internal,
    text: support.text || "",
    start: start ? start.toISOString() : null,
    durationMinutes: support.duration,
    planningType: support.planningType,
  };
}

/** Auswahllisten der Terminmaske: Leistungsarten, Orte, Tickets, Vorgaben. */
export async function options(query, { signal } = {}) {
  const config = settings();
  const companyId = Number(query.companyId) || null;

  const [supportTypes, tickets] = await Promise.all([
    repository().supportTypes({ signal }),
    companyTickets(companyId, signal),
  ]);

  return {
    supportTypes,
    locations: LOCATIONS,
    tickets,
    defaults: {
      supportTypeId: config.defaults.supportTypeId || null,
      location: config.defaults.location,
      internal: config.defaults.internal,
    },
  };
}

/** Die offenen Tickets der Firma - entbehrlich, deshalb still ausfallend. */
async function companyTickets(companyId, signal) {
  if (!companyId) return [];
  try {
    const result = await repository().searchTickets("",
      { companyId, limit: TICKET_LIMIT, signal });
    return result.items.map((ticket) => ({ id: ticket.id, title: ticket.title }));
  } catch {
    return [];
  }
}

/**
 * Die Felder auf einen bestehenden Einsatz zurueckschreiben.
 *
 * Ein geschlossenes Feldobjekt, und zwar mit Absicht: Weder Beginn noch Dauer noch
 * Mitarbeiter noch Planungsart stehen darin. Was nicht im Feldobjekt steht, kann die
 * Maske nicht versehentlich ueberschreiben.
 */
export async function save(supportId, body, { signal } = {}) {
  const fields = readFields(body);
  const result = await repository().saveSupport(supportId, fields, {
    employeeId: session().employeeId(),
    signal,
  });
  return {
    supportId: result.supportId,
    appliedFields: result.applied,
    ignoredFields: result.ignored,
  };
}

/**
 * Einsatz zu einem noch nicht gekoppelten Termin anlegen.
 *
 * Nur in einer Installation ohne Kalendersynchronisation - sonst entstuenden zwei
 * Eintraege je Termin. Die Sperre liegt im Repository und nicht allein hier.
 */
export async function create(body, { signal } = {}) {
  const itemKey = String(body.itemKey || "").trim();
  if (!itemKey) throw new ApiError("APPOINTMENT_NOT_SAVED");

  const start = new Date(body.start);
  if (Number.isNaN(start.getTime())) {
    throw new ApiError("CONFLICT", "Der Beginn des Termins war nicht lesbar.");
  }

  const fields = {
    ...readFields(body),
    uid: body.icalUid || "",
    start,
    durationMinutes: Number(body.durationMinutes) || 0,
    title: body.title || "",
  };

  const result = await repository().createSupport(fields, {
    employeeId: session().employeeId(),
    signal,
  });

  store().rememberAppointmentLink(itemKey, {
    supportId: result.supportId,
    uid: body.icalUid || "",
    startUnix: Math.floor(start.getTime() / 1000),
    source: "created",
  });

  return {
    supportId: result.supportId,
    appliedFields: result.applied,
    ignoredFields: result.ignored,
  };
}

/**
 * Die Felder aus dem Formular.
 *
 * Die Firma ist Pflicht: Jeder Einsatz braucht eine, und TANSS weist ihn sonst ab. Das
 * hier abzufangen erspart dem Techniker eine Fehlermeldung, die nach einem Serverfehler
 * aussieht.
 */
function readFields(body) {
  const companyId = Number(body.companyId) || 0;
  if (companyId <= 0) {
    throw new ApiError("CONFLICT", "Jeder Einsatz braucht eine Firma.");
  }
  const location = String(body.location || "OFFICE");
  return {
    companyId,
    ticketId: Number(body.ticketId) || null,
    typeId: Number(body.typeId) || 0,
    location: LOCATIONS.includes(location) ? location : "OFFICE",
    internal: body.internal !== false,
    text: body.text === undefined || body.text === null ? null : String(body.text),
  };
}
