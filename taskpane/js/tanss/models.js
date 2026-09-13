/**
 * Die Koerper, die an TANSS gehen, und die Zeilen, die aus TANSS kommen.
 *
 * Eine Regel traegt diese ganze Datei: **Nur gesetzte Felder werden gesendet.** TANSS
 * liest ein `null` als "leeren". Ein Schreibkoerper, der jedes nicht ausgefuellte Feld
 * als `null` mitschickt, loescht also genau die Angaben, die der Techniker nicht
 * angefasst hat - und zwar still, mit Erfolgsstatus. `clean()` ist die einzige Stelle,
 * an der ein Koerper entsteht, damit diese Regel nicht an zwoelf Stellen wiederholt
 * werden muss.
 */

/** Der Server schneidet laengere Titel wortlos ab - also schneiden wir sichtbar. */
export const MAX_TICKET_TITLE = 100;

/**
 * Vorschaulaenge eines Ticket-Treffers. Die Vorgabe des Servers schneidet mitten im
 * ersten Satz ab und macht zwei aehnliche Tickets ununterscheidbar.
 */
export const PREVIEW_CHARS = 120;

/**
 * Wirft alles heraus, was nicht gesetzt ist.
 *
 * `undefined`, `null` und `""` gelten als "nicht gesetzt". Die Null und `false` bleiben:
 * Sie sind Angaben, keine Leerstellen - ein `internal: false` muss ankommen.
 */
export function clean(body) {
  const result = {};
  for (const [key, value] of Object.entries(body)) {
    if (value === undefined || value === null || value === "") continue;
    if (Array.isArray(value) && value.length === 0) continue;
    result[key] = value;
  }
  return result;
}

/* ------------------------------------------------------------------- Suche */

/**
 * Koerper von `PUT /api/v1/search`.
 *
 * Zwei Eigenheiten, die man kennen muss: Die Volltextsuche verbindet mehrere Woerter
 * INNERHALB eines Feldes mit UND, nicht ueber Felder hinweg - "Mueller Stoerung" findet
 * also kein Ticket der Firma Mueller mit dem Titel "Stoerung". Und eine leere
 * Trefferliste zusammen mit `tooManyEntries` heisst "Suche verfeinern", nicht "nichts
 * gefunden".
 */
export function searchRequest({ areas, query, configs }) {
  return clean({ areas, query, configs: configs ? clean(configs) : undefined });
}

/** Suchbereich Firmen. */
export function companyConfig({ maxResults }) {
  return clean({ maxResults });
}

/**
 * Suchbereich Ansprechpartner.
 *
 * `inactive: false` ist Absicht und muss gesetzt werden: Ohne die Angabe schliesst
 * TANSS ausgeschiedene Ansprechpartner EIN. Ein ausgeschiedener Melder im Ticket ist
 * ein stiller Fehler, den niemand mehr korrigiert.
 */
export function employeeConfig({ maxResults, companyId }) {
  return { ...clean({ maxResults, companyId }), inactive: false };
}

/** Suchbereich Tickets. Die Vorschaulaenge wird IMMER gesetzt - siehe PREVIEW_CHARS. */
export function ticketConfig({ maxResults, companyId }) {
  return {
    ...clean({ maxResults, companyId }),
    previewContentMaxChars: PREVIEW_CHARS,
  };
}

/** Koerper von `PUT /api/v1/tickets` - die Ticketliste mit Filter. */
export function ticketListFilter({ ids, companies, includeDoneTickets, itemsPerPage, page }) {
  return clean({ ids, companies, includeDoneTickets, itemsPerPage, page });
}

/* ------------------------------------------------------------------ Tickets */

/**
 * Ticketentwurf -> Schreibkoerper.
 *
 * Drei Regeln, die hier und nirgends sonst durchgesetzt werden:
 *
 * - **Titel hart auf 100 Zeichen.** Der Server schneidet wortlos ab; wer sich darauf
 *   verlaesst, bekommt einen Titel, der mitten im Wort endet, und merkt es nie.
 * - **Inhalt als Klartext.** Der Inhalt wird verbatim gespeichert und in der
 *   Oberflaeche in einem Textfeld angezeigt - HTML erschiene dort als Quelltext.
 * - **Zuordnung nur vollstaendig.** Ein `linkId` ohne `linkTypeId` verwirft der Server
 *   stillschweigend.
 *
 * Ein `internal`-Kennzeichen wandert bewusst NICHT mit: TANSS kennt es am Ticket nicht.
 * Das gleichnamige Feld der Maske steuert die Mailablage.
 */
export function ticketWrite(draft) {
  const body = clean({
    companyId: draft.companyId,
    remitterId: draft.remitterId,
    title: String(draft.title || "").trim().slice(0, MAX_TICKET_TITLE),
    content: draft.content || "",
    typeId: draft.typeId,
    statusId: draft.statusId,
    assignedToEmployeeId: draft.assignedToEmployeeId,
    // Der Draht-Name ist `priority`. Ohne Auswahl geht das Feld GAR NICHT mit: Der Server
    // setzt dann seine eigene, je Instanz eingestellte Vorgabe. Eine 0 zu senden waere
    // gleichbedeutend, aber eine hier gewaehlte Zahl waere eine Behauptung ueber eine
    // fremde Einstellung.
    priority: draft.priority,
  });
  if (draft.linkTypeId && draft.linkId) {
    body.linkTypeId = draft.linkTypeId;
    body.linkId = draft.linkId;
  }
  return body;
}

/**
 * Die zulaessigen Stufen der Ticketprioritaet - 1 bis 9, sonst nichts.
 *
 * Die Grenzen stehen hier, weil TANSS sie auf dem Weg ueber die Schnittstelle NICHT
 * prueft: Ein Wert wie 12 oder -1 landete ungeprueft in der Spalte. Geprueft wird er nur
 * auf einem anderen Weg in TANSS, und darauf ist kein Verlass.
 */
export const PRIORITY_MIN = 1;
export const PRIORITY_MAX = 9;

/** Eine Stufe, oder `null`. Alles ausserhalb 1..9 gilt als keine Angabe. */
export function priorityOrNull(value) {
  const stufe = num(value);
  return stufe >= PRIORITY_MIN && stufe <= PRIORITY_MAX ? stufe : null;
}

/**
 * Ansprechpartner -> Schreibkoerper.
 *
 * Der Name kommt als eine Zeichenkette aus dem Mailkopf und wird am LETZTEN Leerzeichen
 * getrennt: "Anna Maria Muster" ergibt Vorname "Anna Maria", Nachname "Muster". Das ist
 * eine Faustregel und in der Maske auch als solche kenntlich - der Benutzer kann beide
 * Felder vor dem Anlegen aendern.
 *
 * Mindestens eine Firmenzuordnung ist Pflicht; ohne sie weist TANSS den Satz ab.
 */
export function employeeWrite({ name, email, companyId }) {
  const clean_name = String(name || "").trim();
  const cut = clean_name.lastIndexOf(" ");
  const first = cut > 0 ? clean_name.slice(0, cut) : "";
  const last = cut > 0 ? clean_name.slice(cut + 1) : clean_name;
  return clean({
    firstName: first,
    lastName: last || clean_name,
    name: clean_name || email,
    emailAddress: email,
    active: true,
    companyAssignments: [{ companyId }],
  });
}

/* ------------------------------------------------------------------- Zeilen */

/** Eine Zahl aus einer Antwort, die auch `null` oder eine Zeichenkette sein darf. */
export function num(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

/** Eine Zeichenkette aus einer Antwort, die auch `null` sein darf. */
export function str(value) {
  return value === null || value === undefined ? "" : String(value);
}

/**
 * Eine einzelne `linkedEntities`-Gruppe als Abbildung `{id: name}`.
 *
 * Ohne diese Aufloesung stuenden in der Trefferliste Firmennummern statt Firmennamen.
 * Die Namen liegen bereits im `meta`-Block DERSELBEN Antwort; sie einzeln nachzuladen
 * kostete einen Zusatzaufruf je Treffer.
 *
 * Der Block ist ein Objekt mit der Kennung als SCHLUESSEL (`{"3": {"name": "..."}}`),
 * nicht eine Liste.
 */
export function linkedNames(meta, entity) {
  const group = ((meta || {}).linkedEntities || {})[entity] || {};
  const resolved = new Map();
  for (const [key, value] of Object.entries(group)) {
    const id = Number.parseInt(key, 10);
    if (!Number.isFinite(id)) continue;
    resolved.set(id, value && typeof value === "object" ? str(value.name) : str(value));
  }
  return resolved;
}

/**
 * Hat der Server gekappt, statt nichts zu finden?
 *
 * TANSS antwortet auf eine zu weite Suche mit einer LEEREN Liste PLUS diesem Hinweis.
 * "Leer" allein heisst "nichts gefunden", "leer plus Hinweis" heisst "Suche
 * verfeinern" - zwei Saetze, die der Benutzer unterscheiden koennen muss.
 *
 * Nachgesehen wird in beiden Behaeltern, die dafuer in Frage kommen. Steht der Hinweis
 * an einer dritten Stelle, ist die Folge eine zu schwache Meldung, kein falsches
 * Ergebnis.
 */
export function tooMany(meta) {
  const containers = [
    (((meta || {}).properties || {}).extras) || {},
    (meta || {}).extras || {},
  ];
  return containers.some(
    (container) => container && typeof container === "object" && Boolean(container.tooManyEntries),
  );
}

/**
 * Eine Firma, wie die Auswahlliste sie zeigt.
 *
 * `inactive` reist mit, damit die Maske eine stillgelegte Firma kennzeichnen kann,
 * statt sie wie jede andere anzubieten - ein Ticket auf eine inaktive Firma ist ein
 * stiller Fehler, den erst die Rechnungsstellung findet.
 */
export function companyRow(item) {
  return {
    id: num(item.id),
    name: str(item.name),
    city: str(item.city),
    postCode: str(item.postCode),
    displayId: str(item.displayId),
    inactive: Boolean(item.inactive),
  };
}

/** Ein Ansprechpartner, wie die Melderauswahl ihn zeigt. */
export function employeeRow(item) {
  return {
    id: num(item.id),
    name: str(item.name),
    email: str(item.emailAddress || item.email),
    departmentId: num(item.departmentId) || null,
    companyIds: companyIdsOf(item),
  };
}

/**
 * Zu welchen Firmen ein Mitarbeiter gehoert.
 *
 * Eine MENGE, kein Feld: In TANSS kann ein Ansprechpartner mehreren Firmen zugeordnet
 * sein - `employeeWrite` schreibt die Zuordnung als `companyAssignments: [{companyId}]`.
 * Wer hier ein einzelnes Feld erwartete, hielte einen Ansprechpartner mit zwei Firmen bei
 * einer davon faelschlich fuer fremd.
 *
 * Gelesen werden alle drei Schreibweisen, die in der Antwort vorkommen koennen. Eine
 * LEERE Menge heisst ausdruecklich "TANSS hat dazu nichts gesagt" und NICHT "gehoert zu
 * keiner Firma" - der Unterschied entscheidet darueber, ob eine Pruefung moeglich ist.
 */
function companyIdsOf(item) {
  const gefunden = new Set();
  const merken = (wert) => {
    const id = num(wert && typeof wert === "object" ? (wert.companyId ?? wert.id) : wert);
    if (id) gefunden.add(id);
  };
  if (Array.isArray(item.companyAssignments)) item.companyAssignments.forEach(merken);
  if (Array.isArray(item.companies)) item.companies.forEach(merken);
  merken(item.companyId);
  return [...gefunden];
}

/**
 * Gehoert diese Zeile nachweislich NICHT zur Firma?
 *
 * Drei Ausgaenge, und der dritte ist der wichtigste:
 *   - die Zeile nennt Firmen und die gesuchte ist dabei  -> gehoert dazu
 *   - die Zeile nennt Firmen und die gesuchte fehlt      -> gehoert nachweislich nicht
 *   - die Zeile nennt gar keine Firma                    -> nicht feststellbar
 *
 * Nur der mittlere Fall rechtfertigt es, jemanden wegzulassen. Den dritten als
 * "gehoert nicht" zu behandeln, leerte die Trefferliste einer Instanz, die das Feld nicht
 * mitliefert - und aus einer Schutzmassnahme wuerde ein Ausfall.
 */
export function belongsToCompany(row, companyId) {
  const ids = Array.isArray(row && row.companyIds) ? row.companyIds : [];
  if (ids.length === 0) return null;
  return ids.includes(num(companyId));
}

/**
 * Ein Ticket, wie Trefferliste und Ticketmaske es zeigen.
 *
 * `companyName` und `statusName` stehen nicht am Ticket, sondern im `linkedEntities`-
 * Block derselben Antwort. Sie werden uebergeben statt nachgeschlagen, damit fuer eine
 * Trefferliste kein einziger Zusatzaufruf entsteht.
 */
export function ticketRow(item, { companies = new Map(), states = new Map(),
  attachedTicketIds = new Set() } = {}) {
  const id = num(item.id);
  const companyId = num(item.companyId);
  const statusId = num(item.statusId);
  return {
    id,
    title: str(item.title),
    preview: str(item.content),
    statusId,
    statusName: states.get(statusId) || "",
    typeId: num(item.typeId),
    companyId,
    companyName: companies.get(companyId) || "",
    assignedToEmployeeId: num(item.assignedToEmployeeId),
    assignedToDepartmentId: num(item.assignedToDepartmentId),
    extTicketId: str(item.extTicketId),
    alreadyHasThisMail: attachedTicketIds.has(id),
  };
}

/** Ein Eintrag einer Auswahlliste. */
export function namedId(item) {
  return { id: num(item.id), name: str(item.name) };
}

/* ------------------------------------------------------------------ Einsaetze */

/** Zuordnungstyp "Firma". Ohne ihn haengt der Einsatz an keiner Firma. */
export const LINK_TYPE_COMPANY = 2;

/** Planungsart, mit der ein selbst angelegter Einsatz entsteht. */
export const CREATE_PLANNING_TYPE = "PLANNED";

/**
 * Ein Einsatz, wie die Terminmaske ihn liest.
 *
 * `date` kommt als Unix-Zeitstempel in SEKUNDEN, `duration` in MINUTEN. Die beiden
 * Einheiten werden im ganzen Pane nur hier angefasst - eine zweite Umrechnungsstelle
 * waere eine Verwechslung, die erst im Kalender auffaellt.
 */
export function supportRow(item) {
  return {
    id: num(item.id),
    ticketId: num(item.ticketId) || null,
    date: num(item.date),
    duration: num(item.duration),
    employeeId: num(item.employeeId),
    companyId: num(item.companyId) || null,
    typeId: num(item.typeId) || null,
    location: str(item.location) || "OFFICE",
    text: str(item.text),
    internal: item.internal !== false,
    planningType: str(item.planningType) || "NONE",
    outlookTitle: str(item.outlookTitle),
    syncGroup: str((item.metaInfos || {}).SYNC_GROUP),
    recurrenceRuleSequenceId: item.recurrenceRuleSequenceId === null
      || item.recurrenceRuleSequenceId === undefined
      ? null
      : num(item.recurrenceRuleSequenceId),
    metaInfos: item.metaInfos && typeof item.metaInfos === "object" ? item.metaInfos : {},
  };
}

/**
 * Ist dieser Eintrag noch ein aenderbarer Termin?
 *
 * Eine gebuchte Leistung, Urlaub oder eine Abwesenheit wird vom Add-in nicht
 * veraendert. Ohne diese Unterscheidung ueberschriebe ein Speichervorgang eine bereits
 * abgerechnete Leistung.
 */
export function isEditableAppointment(support) {
  return support.planningType !== "SUPPORT" && !support.booked;
}

/** Sekunden seit 1970 aus einem Datum. */
export function toUnix(value) {
  const date = value instanceof Date ? value : new Date(value);
  return Math.floor(date.getTime() / 1000);
}

/** Datum aus Sekunden seit 1970 - `null`, wenn nichts gesetzt ist. */
export function fromUnix(seconds) {
  const value = num(seconds);
  return value > 0 ? new Date(value * 1000) : null;
}

/**
 * Feldobjekt -> Schreibkoerper eines Einsatzes.
 *
 * Der Unterschied zwischen Anlegen und Aendern ist ABSICHTLICH gross:
 *
 * - Beim **Aendern** gehen weder Beginn noch Dauer noch Mitarbeiter noch Planungsart
 *   mit. Das Add-in reichert einen Termin um TANSS-Felder an; es verschiebt ihn nicht
 *   und macht aus einem Terminvorschlag keinen festen Termin. Wer den Termin
 *   verschieben will, verschiebt ihn in Outlook.
 * - `metaInfos` ist beim Aendern IMMER abwesend. Der Server ersetzt den kompletten
 *   Satz; ein Teilsatz risse die Kopplung der Kalendersynchronisation.
 * - Die Firma geht immer als Dreiklang: Firmennummer, Zuordnungstyp, Zuordnung.
 *   Die Firmennummer allein erzeugt einen Einsatz OHNE Zuordnung - er saehe richtig
 *   aus und waere in der Firmenansicht trotzdem unsichtbar.
 */
export function supportWrite(fields, { employeeId, forCreate, syncGroup = "" }) {
  const body = {};

  if (fields.companyId) {
    body.companyId = fields.companyId;
    body.linkTypeId = LINK_TYPE_COMPANY;
    body.linkId = fields.companyId;
  }
  if (fields.ticketId) body.ticketId = fields.ticketId;
  if (fields.typeId) body.typeId = fields.typeId;
  body.location = fields.location || "OFFICE";
  body.internal = Boolean(fields.internal);
  if (fields.text !== null && fields.text !== undefined) body.text = fields.text;
  if (fields.title) body.outlookTitle = fields.title;

  if (forCreate) {
    if (!fields.start) throw new Error("Ein neuer Einsatz braucht einen Beginn.");
    body.date = toUnix(fields.start);
    body.duration = fields.durationMinutes;
    body.employeeId = employeeId;
    body.planningType = CREATE_PLANNING_TYPE;
    if (syncGroup) {
      body.metaInfos = { SYNC_GROUP: syncGroup, SYNC_ORIGIN: "OUTLOOK" };
    }
  }
  // Der Aenderungspfad hat hier bewusst KEINEN Zweig. Was nicht gesetzt wird, kann
  // auch nicht gesendet werden - und was nicht gesendet wird, kann nichts loeschen.

  return body;
}

/**
 * Letzte Pruefung unmittelbar vor dem Absenden.
 *
 * Zwei Felder mit Loeschwirkung, an einer Stelle geprueft:
 *
 * - `linkedTechnicians` darf ueberhaupt nicht vorkommen. Mitgesendet legt der Server
 *   Einsaetze fuer hinzugekommene Techniker an und LOESCHT die der entfernten; ein
 *   leeres Feld loescht die Termine aller weiteren Teilnehmer.
 * - `metaInfos` muss eine echte Obermenge des GELESENEN Satzes sein. Der Server
 *   ersetzt den kompletten Satz; alles andere loescht mindestens einen fremden
 *   Eintrag - typischerweise die Kopplung der Kalendersynchronisation.
 *
 * `currentMeta === null` heisst "im selben Vorgang nicht gelesen" und ist ebenfalls ein
 * Verstoss: Die Obermengenregel ist nur so viel wert wie die Aktualitaet des
 * Vergleichssatzes. Beim Anlegen ist der aktuelle Satz per Definition leer - dort wird
 * ausdruecklich ein leeres Objekt uebergeben, nicht `null`.
 */
export function assertWriteSafe(body, { currentMeta }) {
  if (Object.prototype.hasOwnProperty.call(body, "linkedTechnicians")) {
    throw new Error(
      "linkedTechnicians darf niemals gesendet werden: Der Server legt daraufhin "
      + "Einsaetze fuer hinzugekommene Techniker an und loescht die der entfernten.",
    );
  }

  const next = body.metaInfos;
  if (next === undefined || next === null) return;

  if (currentMeta === null || currentMeta === undefined) {
    throw new Error(
      "metaInfos soll geschrieben werden, ohne dass der aktuelle Satz im selben Vorgang "
      + "gelesen wurde. Ohne Vergleichssatz ist nicht zu entscheiden, ob der "
      + "Schreibvorgang fremde Eintraege loescht.",
    );
  }

  const lost = Object.entries(currentMeta)
    .filter(([key, value]) => next[key] !== value)
    .map(([key]) => key)
    .sort();
  if (lost.length > 0) {
    throw new Error(
      "Der Schreibvorgang haette fremde metaInfos geloescht oder ueberschrieben: "
      + `${lost.join(", ")}. Erlaubt ist nur eine Obermenge des gelesenen Satzes.`,
    );
  }
}

/**
 * Vergleicht den gesendeten Koerper mit der Serverantwort.
 *
 * TANSS entfernt Felder, die der Anrufer nicht setzen darf, STILL und antwortet
 * trotzdem mit Erfolg. Ohne diesen Abgleich meldete die Maske "gespeichert", und der
 * Techniker faende beim naechsten Oeffnen die Haelfte seiner Eingaben nicht wieder.
 */
export function appliedFields(sent, returned) {
  const applied = [];
  const ignored = [];
  for (const [key, value] of Object.entries(sent)) {
    if (key === "metaInfos") continue;
    const back = returned[key];
    if (back === undefined) continue;
    if (String(back) === String(value)) applied.push(key);
    else ignored.push(key);
  }
  return { applied, ignored };
}
