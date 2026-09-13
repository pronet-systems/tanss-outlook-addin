/**
 * Suche und Kontaktanlage - was hinter jedem Auswahlfeld der Oberflaeche steckt.
 *
 * Die schreibende Funktion dieses Moduls ist der Ausweg aus der haeufigsten Sackgasse
 * des Arbeitsablaufs: Der Absender steht nicht in TANSS, und ohne ihn verlangt die
 * Firma einen Melder, den es nicht gibt. Ihn hier anzulegen spart den Wechsel in die
 * TANSS-Oberflaeche und den Verlust des halb ausgefuellten Formulars.
 */

import { ApiError } from "../tanss/errors.js";
import { repository, store } from "../runtime.js";

/**
 * Kuerzeste sinnvolle Suchanfrage - dieselbe Grenze, die TANSS selbst zieht: Es zerlegt
 * die Anfrage in Wortmarken und verwirft jede kuerzer als zwei Zeichen. Ein einzelner
 * Buchstabe kostete also einen Aufruf und ergaebe nichts; hier sagen wir das dem
 * Benutzer, statt ihm "keine Treffer" zu zeigen.
 */
const MIN_QUERY_CHARS = 2;

/** Wie viele zuletzt benutzte Tickets die leere Suche mitbringt. */
const RECENT_LIMIT = 10;

function requireQuery(text) {
  const cleaned = String(text || "").trim();
  if (cleaned.length < MIN_QUERY_CHARS) {
    throw new ApiError("CONFLICT", `Bitte mindestens ${MIN_QUERY_CHARS} Zeichen eingeben.`);
  }
  return cleaned;
}

/** Firmensuche hinter dem Firmenwaehler. */
export async function searchCompanies({ q, limit = 25 }, { signal } = {}) {
  return repository().searchCompanies(requireQuery(q), { limit, signal });
}

/** Melderauswahl - die Ansprechpartner der gewaehlten Firma. */
export async function searchEmployees({ q, companyId = null, limit = 25 }, { signal } = {}) {
  return repository().searchEmployees(requireQuery(q), {
    companyId: companyId || null,
    limit,
    signal,
  });
}

/**
 * Ticketsuche fuer "An Ticket anhaengen" und die Ticketzuordnung im Termin.
 *
 * `scope: "all"` hebt die Firmenbindung auf. Das ist kein Rechteloch: TANSS filtert die
 * Suche ohnehin auf das, was dieser Mitarbeiter sehen darf - auf seine sichtbaren
 * Abteilungen und seine Firmen.
 *
 * Die zuletzt benutzten Tickets kommen NUR bei leerer Suche mit. Sie sind der Einstieg
 * in die Maske, nicht ein Anhaengsel jeder Tastatureingabe.
 */
export async function searchTickets({
  q = "", companyId = null, scope = "company", includeDone = false, limit = 25,
  internetMessageId = "",
}, { signal } = {}) {
  const text = String(q || "").trim();
  const scoped = scope === "all" ? null : (companyId || null);
  const attachedTicketIds = store().attachedTicketIds(internetMessageId);

  const result = await repository().searchTickets(text, {
    companyId: scoped,
    limit,
    includeDone,
    attachedTicketIds,
    signal,
  });

  return {
    items: result.items,
    recent: text ? [] : await recentTickets(signal),
    tooMany: result.tooMany,
  };
}

/**
 * Die zuletzt benutzten Tickets, mit frisch geholten Titeln.
 *
 * Gemerkt ist nur die Nummer - die Titel kommen mit EINEM Aufruf nach. So liegt kein
 * Kundeninhalt im Ruhezustand, und ein umbenanntes Ticket heisst hier sofort richtig.
 */
async function recentTickets(signal) {
  const ids = store().recentTickets().slice(0, RECENT_LIMIT);
  if (ids.length === 0) return [];
  try {
    const tickets = await repository().ticketsById(ids, { signal });
    const byId = new Map(tickets.map((ticket) => [ticket.id, ticket]));
    // Reihenfolge des Speichers beibehalten: Sie ist nach "zuletzt benutzt" sortiert,
    // und genau das ist der Wert dieser Liste. Was TANSS nicht mehr kennt, faellt raus.
    return ids
      .filter((id) => byId.has(id))
      .map((id) => ({ id, title: byId.get(id).title }));
  } catch {
    // Entbehrlich: Ohne die Liste fehlt eine Abkuerzung, nicht die Funktion.
    return [];
  }
}

/**
 * Unbekannten Absender als Ansprechpartner anlegen.
 *
 * Eine Telefonnummer legt dieser Weg NICHT an. Verworfen wird sie aber nicht
 * stillschweigend: `ignoredFields` nennt sie, damit niemand die Nummer fuer
 * gespeichert haelt. Zum Abbruch fuehrt sie nicht - das Formular bietet das Feld als
 * optional an, und eine Nummer darin liesse den Vorgang sonst genau in der Sackgasse
 * scheitern, fuer die er existiert.
 */
export async function createContact(companyId, body, { signal } = {}) {
  const mail = String(body.emailAddress || "").trim();
  if (!mail.includes("@") || mail.startsWith("@") || mail.endsWith("@")) {
    throw new ApiError(
      "CONFLICT",
      `„${mail}" ist keine E-Mail-Adresse. Ohne gültige Adresse fände die Melderauswahl `
      + "den Ansprechpartner später nicht wieder.",
    );
  }

  const lastName = String(body.lastName || "").trim();
  if (lastName.length === 0) {
    throw new ApiError("CONFLICT", "Ohne Nachnamen bliebe die Trefferzeile leer.");
  }

  const ignoredFields = String(body.telephoneNumber || "").trim() ? ["telephoneNumber"] : [];
  const name = [String(body.firstName || "").trim(), lastName]
    .filter((part) => part.length > 0)
    .join(" ");

  const contact = await repository().createContact(
    { companyId, name, email: mail },
    { signal },
  );

  return {
    id: contact.id,
    name: contact.name,
    emailAddress: contact.email,
    ignoredFields,
  };
}
