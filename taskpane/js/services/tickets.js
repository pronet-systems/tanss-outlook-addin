/**
 * Ticketmaske und Mailablage - der Kernpfad des Add-ins.
 *
 * Zwei Zusagen tragen dieses Modul, und beide sind aus Schadensfaellen abgeleitet:
 *
 * 1. **Die Feldregeln gelten VOR dem Schreiben.** Ein fehlender Melder wird sonst erst
 *    von TANSS abgelehnt - oder schlimmer: still verworfen, und das Ticket steht ohne
 *    Melder da.
 * 2. **Ab dem Anlegen darf nichts mehr scheitern.** Sobald das Ticket existiert, kennt
 *    nur diese Antwort seine Nummer. Ein durchgereichter Fehler aus der Mailablage
 *    machte aus einem erfolgreichen Vorgang eine Fehlermeldung ohne Nummer - und der
 *    Techniker legte beim zweiten Versuch ein zweites Ticket an.
 */

import { ApiError } from "../tanss/errors.js";
import { priorityOrNull } from "../tanss/models.js";
import { repository, settings, store } from "../runtime.js";
import { titleFromSubject } from "../text.js";

/** Wie viele aehnliche und wie viele offene Tickets die Vorbelegung mitbringt. */
const SIMILAR_LIMIT = 5;
const OPEN_TICKET_LIMIT = 10;

/**
 * Adressen dieser Anbieter gehoeren einer Person, nicht einer Firma.
 *
 * Ohne diese Liste suchte die Firmenaufloesung nach der Domaene eines Freemail-Kontos
 * und faende - bei genuegend Kunden - irgendeine Firma, die zufaellig eine solche
 * Adresse hinterlegt hat. Eine falsch vorbelegte Firma faellt beim schnellen Anlegen
 * niemandem auf, und das Ticket steht danach beim falschen Kunden.
 */
const PUBLIC_MAIL_DOMAINS = new Set([
  "gmail.com", "googlemail.com", "gmx.de", "gmx.net", "gmx.at", "gmx.ch",
  "web.de", "t-online.de", "freenet.de", "outlook.com", "outlook.de",
  "hotmail.com", "hotmail.de", "live.com", "live.de", "icloud.com", "me.com",
  "yahoo.com", "yahoo.de", "aol.com", "posteo.de", "mailbox.org", "protonmail.com",
  "proton.me",
]);

/* ------------------------------------------------------------- Vorbelegung */

/**
 * Die gesamte Vorbelegung der Ticketmaske in EINEM Vorgang.
 *
 * Firma aus Absenderadresse und -domaene, Melder, alle Auswahllisten, aehnliche
 * Tickets, offene Tickets der Firma und die Dublettenwarnung. Betreff und Text baut das
 * Pane parallel dazu aus Office-Bordmitteln - deshalb wartet niemand hierauf, bevor er
 * etwas sieht.
 *
 * Der erste Aufruf ist der Lackmustest: Faellt er aus, ist TANSS nicht erreichbar, und
 * ein halb gefuelltes Formular waere eine Luege. Alle FOLGENDEN Aufrufe sind entbehrlich
 * und fallen still aus - ohne sie fehlt eine Bequemlichkeit, nicht die Funktion.
 */
export async function mailContext(body, { signal } = {}) {
  const address = String(body.senderAddress || "").trim().toLowerCase();

  const { company, candidates, confidence } = await resolveCompany(address, signal);
  const companyId = company ? company.id : null;

  const { remitter, remitterCandidates } = await resolveRemitter(address, companyId, signal);
  const options = await ticketOptions({ companyId }, { signal });
  const title = titleFromSubject(body.subject);

  let similarTickets = [];
  let openTickets = [];
  if (companyId) {
    similarTickets = await quiet(() => similar(companyId, title, signal), []);
    openTickets = await quiet(() => open(companyId, signal), []);
  }

  return {
    company: company ? { ...company, confidence } : null,
    companyCandidates: candidates,
    remitter,
    remitterCandidates,
    title,
    options,
    similarTickets,
    openTickets,
    alreadyAttachedTo: [...store().attachedTicketIds(body.internetMessageId)]
      .map((ticketId) => ({ ticketId, attachedAt: null })),
  };
}

/**
 * Firma aus der Absenderadresse - und wenn das nichts ergibt, aus ihrer Domaene.
 *
 * Uebernommen wird nur ein EINDEUTIGER Treffer. Bei mehreren bleibt die Firma leer und
 * die Maske zeigt die Kandidaten. `confidence` sagt der Oberflaeche, worauf die
 * Vorbelegung beruht - "ueber die Adresse gefunden" ist eine staerkere Aussage als
 * "ueber die Domaene".
 */
async function resolveCompany(address, signal) {
  if (!address) return { company: null, candidates: [], confidence: "none" };

  const hits = (await repository().searchCompanies(address, { limit: 25, signal })).items;
  const exact = hits.filter((item) => item.email && item.email.toLowerCase() === address);
  if (exact.length === 1) return { company: exact[0], candidates: dedupe(hits), confidence: "address" };
  if (hits.length === 1) return { company: hits[0], candidates: dedupe(hits), confidence: "address" };

  const domain = address.slice(address.lastIndexOf("@") + 1);
  if (!domain || PUBLIC_MAIL_DOMAINS.has(domain)) {
    return { company: null, candidates: dedupe(hits), confidence: "none" };
  }

  const byDomain = await quiet(
    async () => (await repository().searchCompanies(domain, { limit: 25, signal })).items,
    [],
  );
  const merged = dedupe([...hits, ...byDomain]);
  if (hits.length === 0 && byDomain.length === 1) {
    return { company: byDomain[0], candidates: merged, confidence: "domain" };
  }
  return { company: null, candidates: merged, confidence: "none" };
}

/**
 * Melder aus der Absenderadresse.
 *
 * Gesucht wird ueber dieselbe Flaeche wie fuer die Melderauswahl. Uebernommen wird nur
 * ein eindeutiger Treffer; bei mehreren zeigt die Maske die Kandidaten, statt einen
 * davon zu raten.
 */
async function resolveRemitter(address, companyId, signal) {
  if (!address) return { remitter: null, remitterCandidates: [] };

  const found = await quiet(
    async () => (await repository().searchEmployees(address, { companyId, limit: 25, signal })).items,
    [],
  );

  const exact = found.filter((item) => item.email && item.email.toLowerCase() === address);
  if (exact.length === 1) return { remitter: exact[0], remitterCandidates: found };
  if (found.length === 1) return { remitter: found[0], remitterCandidates: found };
  return { remitter: null, remitterCandidates: found };
}

/** Tickets derselben Firma mit aehnlichem Titel - gegen das doppelt angelegte Ticket. */
async function similar(companyId, title, signal) {
  if (String(title || "").trim().length < 3) return [];
  const result = await repository().searchTickets(title,
    { companyId, limit: SIMILAR_LIMIT, signal });
  return result.items;
}

/** Die offenen Tickets der Firma - der haeufigste Fall ist "gehoert ans bestehende". */
async function open(companyId, signal) {
  const result = await repository().searchTickets("",
    { companyId, limit: OPEN_TICKET_LIMIT, signal });
  return result.items;
}

/** Kandidatenliste ohne Wiederholungen, Reihenfolge der ersten Nennung. */
function dedupe(companies) {
  const seen = new Map();
  for (const company of companies) {
    if (!seen.has(company.id)) seen.set(company.id, company);
  }
  return [...seen.values()];
}

/** Fuehrt etwas Entbehrliches aus. Faellt es aus, gilt der Ersatzwert. */
async function quiet(work, fallback) {
  try {
    return await work();
  } catch {
    return fallback;
  }
}

/* ----------------------------------------------------------- Auswahllisten */

/**
 * Alle Auswahllisten und Feldregeln in einem Vorgang.
 *
 * Wird bei jedem Wechsel von Firma, Typ oder Zuweisung neu geholt: TANSS macht die
 * zulaessigen Werte voneinander abhaengig, und eine Liste, die nach dem Firmenwechsel
 * stehen bleibt, bietet Werte an, die der Server anschliessend still verwirft.
 */
export async function ticketOptions(query, { signal } = {}) {
  return repository().ticketOptions({ companyId: query.companyId || null, signal });
}

/* -------------------------------------------------------------- Ticketanlage */

/**
 * Ticket aus der gelesenen Nachricht anlegen UND die Nachricht im selben Vorgang
 * ablegen.
 *
 * Die Antwort ist immer aussagekraeftig: `mail.status` sagt, was aus der Nachricht
 * geworden ist, und `applied` spiegelt die SERVERANTWORT - TANSS entfernt Felder, die
 * der Anrufer nicht setzen darf, still und antwortet trotzdem mit Erfolg. Wer aus dem
 * Statuscode auf die Wirkung schloesse, meldete dem Techniker eine Zuweisung, die es
 * nicht gibt.
 */
export async function createTicket(form, { signal } = {}) {
  // OHNE MAIL KEIN TICKET, und die Regel steht HIER und nicht nur in der Maske. Das
  // Werkzeug ist dazu da, eine E-Mail in TANSS zu bringen; ein Ticket ohne sie waere ein
  // leerer Vorgang, den jemand spaeter von Hand nachziehen muesste - und bis dahin sieht
  // es aus, als sei alles erledigt. Eine Maske, die das kuenftig anders sieht, kommt hier
  // nicht vorbei.
  if (!form.get("eml")) {
    throw new ApiError("CONFLICT", "Ohne die E-Mail wird kein Ticket angelegt.");
  }

  const draft = readTicket(form);
  const options = await repository().ticketOptions({
    companyId: draft.companyId || null,
    typeId: draft.typeId || null,
    assigneeId: draft.assignedToEmployeeId || null,
    signal,
  });
  validate(draft, options);

  const ticket = await repository().createTicket(draft, { signal });

  // AB HIER EXISTIERT DAS TICKET. Kein Fehler darf den Vorgang noch scheitern lassen.
  store().touchTicket(ticket.id);
  const mail = await attachOrReport(ticket.id, form, signal);

  return {
    ticketId: ticket.id,
    url: ticketUrl(ticket.id),
    applied: {
      companyId: ticket.companyId,
      statusId: ticket.statusId,
      typeId: ticket.typeId,
      assignedToEmployeeId: ticket.assignedToEmployeeId,
      assignedToDepartmentId: ticket.assignedToDepartmentId,
      extTicketId: ticket.extTicketId,
    },
    mail,
  };
}

/** Nachricht an ein BESTEHENDES Ticket haengen. */
export async function attachMail(ticketId, form, { signal } = {}) {
  const result = await attach(ticketId, form, signal);
  store().touchTicket(ticketId);
  return { status: "attached", ...result };
}

/**
 * Die Zustaende, die TANSS je hochgeladener Datei meldet - auf eigene Fehlercodes.
 *
 * Was hier nicht steht, faellt auf INTERNAL zurueck. Das ist Absicht: Ein neuer Zustand
 * soll als unerwarteter Fehler auffallen und nicht stillschweigend als Erfolg gelten.
 */
const MAIL_STATUS_CODES = {
  TOO_BIG: "MAIL_TOO_LARGE",
  EMPTY: "MAIL_EMPTY",
};

/**
 * Die Mailablage innerhalb der Ticketanlage - sie darf NIE werfen.
 *
 * `code` traegt den Sammelbegriff, den die Erfolgsseite als Ueberschrift benutzt, und
 * `cause` die genaue Ursache. Die entscheidet, ob "Mail jetzt anhaengen" ueberhaupt
 * etwas nuetzt: Bei einer zu grossen Nachricht hilft kein zweiter Versuch, bei einem
 * nicht erreichbaren TANSS sehr wohl.
 */
async function attachOrReport(ticketId, form, signal) {
  try {
    return { status: "attached", ...(await attach(ticketId, form, signal)) };
  } catch (error) {
    return {
      status: "failed",
      code: "TICKET_CREATED_MAIL_FAILED",
      cause: error instanceof ApiError ? error.code : "INTERNAL",
      message: error && error.message ? error.message : "",
    };
  }
}

/**
 * Pruefen, ablegen, gegenpruefen, festhalten.
 *
 * Die Dublettenpruefung ist eine WARNUNG, keine Sperre: Es gibt berechtigte Gruende,
 * dieselbe Nachricht erneut abzulegen, und `force` laesst sie zu. TANSS selbst prueft
 * auf keinem Anhaengepfad, ob eine Nachricht schon am Ticket haengt.
 */
async function attach(ticketId, form, signal) {
  const blob = form.get("eml");
  if (!blob) {
    throw new ApiError("CONFLICT", "Es wurde keine Nachricht mitgeschickt.");
  }

  const messageId = String(form.get("internetMessageId") || "");
  const force = String(form.get("force") || "") === "true";
  if (!force && messageId && store().attachedTicketIds(messageId).has(ticketId)) {
    throw new ApiError("MAIL_ALREADY_ATTACHED");
  }

  const maxBytes = settings().limits.maxEmlBytes;
  if (blob.size > maxBytes) {
    throw new ApiError("MAIL_TOO_LARGE");
  }

  const subject = String(form.get("subject") || "");
  const outcome = await repository().uploadEml(ticketId, {
    filename: blob.name || "nachricht.eml",
    blob,
    description: subject.slice(0, 200) || "Aus Outlook abgelegt",
  }, { signal });

  if (outcome.status !== "OK") {
    // TANSS kennt vier Zustaende: OK, TOO_BIG, ERROR und EMPTY. Die letzten beiden fielen
    // frueher in denselben Sammelfall. "EMPTY" nennt aber eine Ursache, die der Techniker
    // sofort versteht und selbst beheben kann - sie gehoert nicht hinter "unerwarteter
    // Fehler".
    throw new ApiError(MAIL_STATUS_CODES[outcome.status] || "INTERNAL");
  }
  if (outcome.storedAsDocument) {
    throw new ApiError("MAIL_STORED_AS_DOCUMENT");
  }

  // Erst jetzt festhalten, und nur bei echtem Erfolg. Ein Vermerk ueber eine
  // Nachricht, die gar nicht am Ticket haengt, unterdrueckte spaeter den zweiten,
  // richtigen Versuch.
  if (messageId) store().rememberAttached(messageId, ticketId);

  return {
    tanssMailId: outcome.tanssMailId,
    storedAsDocument: false,
    bytes: outcome.sizeBytes,
    // Woher die Bestaetigung stammt, wird durchgereicht und nicht unterwegs verschluckt:
    // Ein Erfolg, den die Gegenprobe am Ticket festgestellt hat, ist ein Erfolg - aber
    // nicht derselbe wie eine Zusage des Servers, und die Erfolgsseite sagt das.
    confirmedBy: outcome.confirmedBy || "response",
  };
}

/**
 * Die Formularfelder des Tickets.
 *
 * Das Tripel aus Zuordnungstyp und Zuordnungskennung setzt DIESE Stelle aus der
 * Firmennummer - es ist eine TANSS-Eigenheit und gehoert nicht in eine Oberflaeche.
 */
function readTicket(form) {
  let body;
  try {
    body = JSON.parse(String(form.get("ticket") || "{}"));
  } catch {
    throw new ApiError("CONFLICT", "Die Ticketangaben waren nicht lesbar.");
  }

  const draft = {
    companyId: Number(body.companyId) || 0,
    remitterId: Number(body.remitterId) || 0,
    title: String(body.title || "").trim(),
    content: String(body.content || ""),
    typeId: Number(body.typeId) || 0,
    statusId: Number(body.statusId) || 0,
    assignedToEmployeeId: Number(body.assignedToEmployeeId) || 0,
    // `null` heisst: nicht mitsenden. Dann setzt TANSS seine eigene Vorgabe - die ist je
    // Instanz eingestellt und gehoert nicht hierher.
    priority: priorityOrNull(body.priority),
  };
  if (draft.companyId) {
    draft.linkTypeId = 2;
    draft.linkId = draft.companyId;
  }
  return draft;
}

/**
 * Die Feldregeln der Firma, geprueft VOR dem Schreiben.
 *
 * TANSS meldet sie je Firma; sie erst beim Schreiben wirken zu lassen hiesse, ein
 * abgelehntes oder - schlimmer - ein still beschnittenes Ticket zu erzeugen.
 */
function validate(draft, options) {
  if (!draft.title) {
    throw new ApiError("TITLE_REQUIRED");
  }
  if (options.remitterRequired && !draft.remitterId) {
    throw new ApiError("REMITTER_REQUIRED");
  }
  if (options.forceAssignment && !draft.assignedToEmployeeId) {
    throw new ApiError("CONFLICT", "Für diese Firma ist eine Zuweisung Pflicht.");
  }
}

/**
 * Die TANSS-Weboberflaeche zu einem Ticket.
 *
 * Die einzige Stelle, an der eine Ticketadresse entsteht. Die Nummer wird kodiert
 * eingesetzt; eine unkodiert zusammengesetzte Adresse braeche erst beim Anklicken, also
 * weit entfernt von der Stelle, an der sie entstanden ist.
 */
export function ticketUrl(ticketId) {
  const base = String(settings().frontendBase || "").replace(/\/+$/, "");
  if (!base) return "";
  const query = new URLSearchParams({
    section: "bug",
    sub: "view",
    bugID: String(Number(ticketId) || 0),
  });
  return `${base}/index.php?${query}`;
}
