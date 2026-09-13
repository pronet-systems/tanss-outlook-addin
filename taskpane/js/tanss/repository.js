/**
 * Die fachlichen TANSS-Operationen des Add-ins.
 *
 * Diese Schicht kennt TANSS und sonst nichts: kein Office, keine Oberflaeche, keinen
 * oertlichen Speicher. Was sie zurueckgibt, sind fertige Zeilen fuer die Maske - und
 * zwar samt der Namen aus dem `linkedEntities`-Block derselben Antwort, damit eine
 * Trefferliste keinen einzigen Zusatzaufruf kostet.
 *
 * Zwei Entscheidungen, die sich durchziehen:
 *
 * - **Ein Ausfall ist kein Abbruch, wo er entbehrlich ist.** Faellt eine Auswahlliste
 *   aus, bleibt sie leer und die Maske sagt es. Faellt die Ticketanlage aus, bricht der
 *   Vorgang ab. Der Unterschied steht an jeder Stelle einzeln begruendet.
 * - **Ein Erfolg ohne Beleg ist kein Erfolg.** Antwortet TANSS auf eine Anlage mit 201,
 *   aber ohne Nummer, wird das als Fehlschlag gemeldet. Alles andere hiesse, dem
 *   Techniker einen Vorgang zu bestaetigen, von dem niemand weiss, ob er stattfand.
 */

import { ApiError, reasonOf } from "./errors.js";
import { canonicalUid, formatSyncGroup, syncGroupMatches } from "./uid.js";
import {
  appliedFields,
  assertWriteSafe,
  belongsToCompany,
  companyConfig,
  companyRow,
  employeeConfig,
  employeeRow,
  employeeWrite,
  linkedNames,
  namedId,
  num,
  searchRequest,
  str,
  ticketConfig,
  ticketListFilter,
  ticketRow,
  ticketWrite,
  tooMany,
  fromUnix,
  isEditableAppointment,
  supportRow,
  supportWrite,
  toUnix,
} from "./models.js";

/**
 * Die Technikerliste aendert sich im Tagesbetrieb nicht. Eine Viertelstunde spart
 * hunderte Abrufe und verzoegert einen neuen Kollegen um hoechstens diese Zeit.
 */
const TECHNICIAN_CACHE_MS = 900_000;

/** Die Auswahllisten der Ticketmaske haengen an der Firma - deshalb der Schluessel. */
const OPTIONS_CACHE_MS = 600_000;

/**
 * Wie weit Beginn und Dauer abweichen duerfen, damit ein Einsatz als derselbe Termin
 * gilt. Zwei Minuten decken Rundungen und Uhrenversatz ab; mehr machte aus dem Abgleich
 * eine Ratefunktion.
 */
const MATCH_WINDOW_MINUTES = 2;

/**
 * Zeitraum, den die Terminaufloesung um den Termin herum abfragt. Weit genug, dass ein
 * in TANSS verschobener Einsatz ueber seine Kopplung noch gefunden wird, eng genug, dass
 * nicht der halbe Kalender uebertragen wird.
 */
const SUPPORT_SPAN_MS = 86_400_000;

/** Fenster, in dem ein frisch angelegtes Dokument als "gerade entstanden" gilt. */
const DOCUMENT_WINDOW_SECONDS = 60;

/** Gegenlaeufige Uhren zwischen Arbeitsplatz und Server. */
const CLOCK_SLACK_SECONDS = 5;

/**
 * So benennt TANSS eine Nachricht, die es NICHT als Mail lesen konnte und deshalb als
 * gewoehnliche Datei ablegt. In der Mailansicht des Tickets ist sie dann unsichtbar -
 * genau das ist der Fall, den die Gegenprobe finden soll.
 */
const EML_DOCUMENT = /^eml_(\d+)_(\d+)_/;

/** Ein Zwischenspeicher mit Verfallszeit. Mehr braucht diese Schicht nicht. */
class Cache {
  constructor(ttlMs) {
    this.ttlMs = ttlMs;
    this.entries = new Map();
  }

  get(key) {
    const entry = this.entries.get(key);
    if (!entry) return null;
    if (entry.until < Date.now()) {
      this.entries.delete(key);
      return null;
    }
    return entry.value;
  }

  put(key, value) {
    this.entries.set(key, { value, until: Date.now() + this.ttlMs });
  }

  clear() {
    this.entries.clear();
  }
}

export class TanssRepository {
  /**
   * @param {{client: import("./client.js").TanssClient, config: object}} options
   */
  constructor({ client, config }) {
    this.client = client;
    this.config = config;
    this._technicians = new Cache(TECHNICIAN_CACHE_MS);
    this._departments = new Cache(TECHNICIAN_CACHE_MS);
    this._options = new Cache(OPTIONS_CACHE_MS);
  }

  /* ---------------------------------------------------------------- Identitaet */

  /**
   * Alle Techniker - Nummer und Name.
   *
   * Frueher lief das ueber die Integrationsflaeche `/api/tanss.x/v1/technicians`, weil
   * jene zusaetzlich die Adresse liefert. Diese Flaeche verlangt aber eine ANDERE
   * Tokenart als eine gewoehnliche Anmeldung: Mit dem Token eines angemeldeten Technikers
   * ist sie grundsaetzlich nicht erreichbar, nicht nur gelegentlich. Das Pane meldete
   * deshalb zuverlaessig "Die Technikerliste war nicht abrufbar".
   *
   * Die Adresse, um derentwillen dieser Weg gewaehlt wurde, braucht niemand: Der einzige
   * Aufrufer zeigt den eigenen Namen an und vergleicht dafuer Kennungen. Diese Route
   * liegt unter derselben Rolle wie der uebrige Fachzugriff und liefert, was gebraucht
   * wird.
   *
   * Sie setzt voraus, dass der Angemeldete in TANSS als Techniker oder Freiberufler
   * gefuehrt wird. Ist er das nicht, faellt sie aus - und der Grund steht in der Meldung.
   */
  async listTechnicians() {
    const cached = this._technicians.get("all");
    if (cached) return cached;
    const payload = (await this.client.get("/api/v1/employees/technicians")) || [];
    const list = (Array.isArray(payload) ? payload : [])
      .filter((item) => item && item.id)
      .map((item) => ({
        id: num(item.id),
        name: str(item.name),
        email: str(item.emailAddress || item.email).toLowerCase(),
      }));
    this._technicians.put("all", list);
    return list;
  }

  /**
   * Die Abteilungen, denen ein Techniker angehoert.
   *
   * In TANSS kann ein Mitarbeiter MEHREREN Abteilungen angehoeren: Neben dem Feld am
   * Mitarbeiter selbst gibt es eine eigene Zuordnung. Gelesen werden deshalb beide
   * Schreibweisen, und das Ergebnis ist eine Menge - wer nur das eine Feld naehme, boete
   * einem Techniker mit zwei Abteilungen nur eine an.
   *
   * Geholt wird ueber die Einzelabfrage am Mitarbeiter. Die naheliegendere Route
   * `employees/{id}/departments` liegt auf der ERP-Flaeche und verlangt eine Rolle, die
   * ein Techniker nicht hat - sie schiede mit 403 aus.
   *
   * Faellt der Abruf aus, ist das Ergebnis LEER und nicht etwa die vollstaendige Liste.
   * Ein Rueckfall auf alle Abteilungen waere genau der Zustand, der abgestellt werden
   * soll, nur ohne Hinweis darauf.
   */
  async employeeDepartments(employeeId, { signal, failures = null } = {}) {
    const id = num(employeeId);
    if (!id) return [];
    const cached = this._departments.get(String(id));
    if (cached) return cached;

    const pfad = `/api/v1/employees/${id}`;
    try {
      const payload = (await this.client.get(pfad, { signal })) || {};
      const gefunden = departmentIdsOf(payload);
      this._departments.put(String(id), gefunden);
      return gefunden;
    } catch (error) {
      if (failures) failures.push({ path: pfad, reason: reasonOf(error) });
      return [];
    }
  }

  /**
   * Eigener Zustand, unter anderem die eigene Firma.
   *
   * Entbehrlich und muss es bleiben: Die eigene Firmennummer fehlt in der Antwort,
   * sobald der Handelnde in TANSS als Kunde gefuehrt wird.
   */
  async ownState() {
    const payload = await this.client.get("/api/v1/employees/ownState");
    return payload && typeof payload === "object" ? payload : {};
  }

  /* ---------------------------------------------------------------- Suche */

  /** Firmensuche hinter dem Firmenwaehler. */
  async searchCompanies(query, { limit = 25, signal } = {}) {
    const { content, meta } = await this.client.put(
      "/api/v1/search",
      searchRequest({
        areas: ["COMPANY"],
        query,
        configs: { company: companyConfig({ maxResults: limit }) },
      }),
      { retry: true, wantMeta: true, signal },
    );
    const items = ((content || {}).companies || []).map(companyRow);
    return { items, tooMany: tooMany(meta) };
  }

  /**
   * Melderauswahl - die Ansprechpartner der gewaehlten Firma.
   *
   * Die Suche ist der einzige Weg an die Adresse eines Ansprechpartners; die
   * Mitarbeiterliste einer Firma fuehrt sie nicht.
   */
  async searchEmployees(query, { companyId = null, limit = 25, signal } = {}) {
    const { content, meta } = await this.client.put(
      "/api/v1/search",
      searchRequest({
        areas: ["EMPLOYEE"],
        query,
        configs: { employee: employeeConfig({ maxResults: limit, companyId }) },
      }),
      { retry: true, wantMeta: true, signal },
    );
    const alle = ((content || {}).employees || []).map(employeeRow);

    // Die Sperre steht HIER und nicht in der Maske. Das ist die einzige Tuer, durch die
    // ein Ansprechpartner in einen Ticketentwurf gelangt; was hier nicht herauskommt,
    // kann keine Oberflaeche anbieten - auch keine, die erst spaeter gebaut wird.
    //
    // Dass die Anfrage die Firma bereits mitgibt, genuegt nicht: Ob TANSS danach filtert,
    // ist eine Zusage des Servers, und eine ungepruefte Zusage ist keine. Ein Ticket am
    // falschen Kunden mit dem Namen einer fremden Person ist kein Schoenheitsfehler.
    const gefragt = num(companyId) || null;
    const items = gefragt
      ? alle.filter((row) => belongsToCompany(row, gefragt) !== false)
      : alle;

    // Ehrlich ueber die Grenze der Pruefung: Nennt eine gelieferte Zeile ueberhaupt keine
    // Firma, ist die Zugehoerigkeit nicht feststellbar. Dann wird sie gezeigt - sonst
    // braeche die Suche auf einer Instanz, die das Feld nicht mitliefert - aber die Maske
    // sagt es, statt eine Pruefung zu behaupten, die nicht stattgefunden hat.
    const unverified = gefragt !== null
      && items.some((row) => belongsToCompany(row, gefragt) === null);

    return { items, tooMany: tooMany(meta), unverified };
  }

  /**
   * Ticketsuche.
   *
   * Drei Eingaben, drei Wege:
   *
   * - **Nur Ziffern** - das ist eine Ticketnummer und keine Suche. Die Volltextsuche
   *   faende die Nummer sonst nur, wenn sie zufaellig im Text vorkommt.
   * - **Leer** - die offenen Tickets der Firma.
   * - **Sonst** - die Volltextsuche.
   */
  async searchTickets(query, {
    companyId = null, limit = 25, includeDone = false,
    attachedTicketIds = new Set(), signal,
  } = {}) {
    const text = String(query || "").trim();

    if (/^\d+$/.test(text)) {
      const items = await this.ticketsById([Number.parseInt(text, 10)],
        { attachedTicketIds, signal });
      return { items, tooMany: false };
    }
    if (!text) {
      const items = await this._ticketsOfCompany(companyId,
        { limit, includeDone, attachedTicketIds, signal });
      return { items, tooMany: false };
    }

    const { content, meta } = await this.client.put(
      "/api/v1/search",
      searchRequest({
        areas: ["TICKET"],
        query: text,
        configs: { ticket: ticketConfig({ maxResults: limit, companyId }) },
      }),
      { retry: true, wantMeta: true, signal },
    );
    return {
      items: this._rows((content || {}).tickets || [], meta, attachedTicketIds),
      tooMany: tooMany(meta),
    };
  }

  /** Tickets zu bekannten Nummern - ein Aufruf fuer beliebig viele. */
  async ticketsById(ids, { attachedTicketIds = new Set(), signal } = {}) {
    if (!ids || ids.length === 0) return [];
    const { content, meta } = await this.client.put(
      "/api/v1/tickets",
      ticketListFilter({ ids }),
      { retry: true, wantMeta: true, signal },
    );
    return this._rows(content || [], meta, attachedTicketIds);
  }

  async _ticketsOfCompany(companyId, { limit, includeDone, attachedTicketIds, signal }) {
    if (!companyId) return [];
    const { content, meta } = await this.client.put(
      "/api/v1/tickets",
      ticketListFilter({
        companies: [companyId],
        includeDoneTickets: includeDone,
        itemsPerPage: limit,
        page: 1,
      }),
      { retry: true, wantMeta: true, signal },
    );
    return this._rows(content || [], meta, attachedTicketIds);
  }

  _rows(tickets, meta, attachedTicketIds) {
    const companies = linkedNames(meta, "companies");
    const states = linkedNames(meta, "ticketStates");
    return (Array.isArray(tickets) ? tickets : [])
      .map((item) => ticketRow(item, { companies, states, attachedTicketIds }));
  }

  /* ---------------------------------------------------------------- Ticketmaske */

  /**
   * Alle Auswahllisten und Feldregeln der Ticketmaske in EINEM Vorgang.
   *
   * Die Listen kommen aus den dokumentierten Routen und sind damit UNGEFILTERT;
   * `source` sagt das der Maske, statt eine Vollstaendigkeit vorzutaeuschen, die sie
   * nicht hat.
   *
   * Faellt eine Liste aus, bleibt sie leer - die Maske zeigt das Feld dann ohne
   * Auswahl. Das ist die mildere Annahme: Ein erfundener Zwang hielte den Techniker
   * von einem gueltigen Ticket ab.
   */
  async ticketOptions({ companyId = null, typeId = null, assigneeId = null,
    departmentId = null, signal } = {}) {
    const key = `${companyId || 0}|${typeId || 0}|${assigneeId || 0}|${departmentId || 0}`;
    const cached = this._options.get(key);
    if (cached) return cached;

    // Was auf dem Weg ausfaellt, wird hier gesammelt statt verschluckt. Die Liste geht
    // mit in die Antwort: Der Techniker sieht, DASS etwas fehlt, und der Administrator
    // sieht, WARUM - ohne dass jemand ein Protokoll aufmachen muss.
    const failures = [];

    const rules = TanssRepository.FIELD_RULES;
    const [types, states, technicians, departments] = await Promise.all([
      // Neben den Ticketstatus, und aus demselben Grund erreichbar: Das "admin" im Pfad
      // ist ein Namensteil, keine Rollenforderung - diese Flaeche liegt unter derselben
      // Rolle wie der uebrige Fachzugriff. Es gibt eine zweite, dokumentierte Typenliste
      // auf der ERP-Flaeche; die verlangt eine Rolle, die ein Techniker nicht hat, und
      // waere mit seinem Token immer abgewiesen worden.
      //
      // Ein `rank` fuehrt die Liste nicht; die Reihenfolge kommt vom Server und wird nicht
      // umsortiert. Stillgelegte Typen werden nicht angeboten.
      //
      // Die Typen gehoeren NICHT in das Manifest: Ein dort eingetragener Typ altert - ein
      // in TANSS neu angelegter oder umbenannter kaeme nie an, ohne dass jemand ein
      // Manifest erzeugt, die Version erhoeht und neu ausrollt.
      this._list("/api/v1/admin/ticketTypes", signal, (raw) => raw
        .filter((item) => item && item.id && item.active !== false)
        .map(namedId), failures),
      this._list("/api/v1/admin/ticketStates", signal, (raw) => raw
        .filter((item) => item && item.id && item.active !== false)
        .sort((a, b) => num(a.rank) - num(b.rank))
        .map(namedId), failures),
      this._list("/api/v1/employees/technicians", signal,
        (raw) => raw.filter((item) => item && item.id).map(namedId), failures),
      this._list("/api/v1/employees/departments", signal,
        (raw) => raw.filter((item) => item && item.id).map(namedId), failures),
    ]);

    const options = {
      // Die API sticht die Konfiguration. Der Schluessel in der config.json bleibt als
      // Rueckfall fuer eine Installation, deren TANSS die Adminroute nicht fuehrt - genau
      // so, wie `defaults.supportTypeId` fuer die Leistungsarten einspringt. Auf der
      // gemeinsam genutzten Ablage ist er leer und bleibt es; dort traegt allein die API.
      types: types.length > 0 ? types : (this.config.ticketTypes || []),
      states,
      technicians,
      departments,
      ...rules,
      failures,
    };
    this._options.put(key, options);
    return options;
  }

  /**
   * Die Feldregeln - es gibt sie fuer diesen Zweck nicht, und das wird gesagt.
   *
   * Frueher stand hier ein Aufruf von `GET /api/v1/tickets/`, aus dessen meta-Block die
   * Kennzeichen `remitterIsAMandatoryField` und `forceAssignment` gelesen wurden. Diese
   * Route gibt es nicht: Unter `/api/v1/tickets` sind nur POST und PUT definiert. Der
   * Aufruf kostete bei jedem Oeffnen der Maske eine Umlaufzeit und konnte nie etwas
   * liefern - die Maske meldete deshalb IMMER, die Listen staemmten aus dem Rueckfall.
   *
   * Es gibt Routen, die diese Kennzeichen fuehren. Ihr Anfragekoerper kennt aber nur eine
   * Feldauswahl und einen Zusammenhang, KEINE Firma - eine Feldsteuerung je Firma, wie
   * die Ticketmaske sie braeuchte, laesst sich daraus nicht holen. Der Aufruf entfaellt
   * deshalb ersatzlos.
   *
   * Was bleibt, ist die mildere Annahme: kein Melderzwang, keine Zuweisungspflicht. Sie
   * ist die richtige - ein erfundener Zwang hielte den Techniker von einem gueltigen
   * Ticket ab, waehrend eine fehlende Pflicht TANSS beim Anlegen selbst auffaellt.
   *
   * `autoAssignedEmployeeId` ist ersatzlos gestrichen: Das Feld kommt in TANSS nicht vor.
   * Die Maske hat daraus eine Vorbelegung gelesen, die niemand je gesendet hat.
   */
  static get FIELD_RULES() {
    return { remitterRequired: false, forceAssignment: false };
  }

  /**
   * Eine Auswahlliste, die ausfallen darf. Faellt sie aus, bleibt sie leer.
   *
   * Entbehrlich heisst nicht unbemerkt: Der Grund wandert nach `failures`, statt im
   * `catch` zu verschwinden. Ohne ihn steht in der Maske "war nicht abrufbar", und wer
   * das untersuchen soll, faengt bei null an - obwohl der Grund im Augenblick des
   * Scheiterns vorlag.
   */
  async _list(path, signal, shape, failures = null) {
    try {
      const raw = (await this.client.get(path, { signal })) || [];
      return shape(Array.isArray(raw) ? raw : []);
    } catch (error) {
      if (failures) failures.push({ path, reason: reasonOf(error) });
      return [];
    }
  }

  /**
   * Leistungsarten fuer die Terminmaske.
   *
   * Faellt die Route aus, bleibt die konfigurierte Vorgabe. Eine leere Auswahlliste
   * waere schlimmer - der Benutzer koennte den Termin dann gar nicht speichern.
   */
  async supportTypes({ signal } = {}) {
    const types = await this._list("/api/v1/supportTypes/active", signal,
      (raw) => raw.filter((item) => item && item.id).map(namedId));
    if (types.length > 0) return types;
    const fallback = num(this.config.defaults && this.config.defaults.supportTypeId);
    return fallback ? [{ id: fallback, name: "Vorgabe" }] : [];
  }

  /* ---------------------------------------------------------------- Anlegen */

  /**
   * Ticket anlegen.
   *
   * Die zurueckgegebene Zeile spiegelt die SERVERANTWORT, nicht den Entwurf: TANSS
   * entfernt Felder, die der Anrufer nicht setzen darf, still und antwortet trotzdem
   * mit Erfolg. Wer aus dem Statuscode auf die Wirkung schloesse, meldete dem Techniker
   * eine Zuweisung, die es nicht gibt.
   */
  async createTicket(draft, { signal } = {}) {
    const { content, meta } = await this.client.post(
      "/api/v1/tickets", ticketWrite(draft), { wantMeta: true, signal });
    const ticket = content || {};
    if (!num(ticket.id)) {
      throw new ApiError(
        "INTERNAL",
        "TANSS hat auf die Ticketanlage mit Erfolg, aber ohne Ticketnummer geantwortet. "
        + "Ob ein Ticket entstanden ist, laesst sich so nicht sagen - der Vorgang wird "
        + "deshalb als fehlgeschlagen gemeldet.",
      );
    }
    return this._rows([ticket], meta, new Set())[0];
  }

  /**
   * Unbekannten Absender als Ansprechpartner anlegen.
   *
   * Die dokumentierte Antwort ist als Liste beschrieben, in der Praxis kommt auch ein
   * einzelnes Objekt vor. Beides wird gelesen - ein Formatunterschied darf keinen
   * angelegten Kontakt verschlucken.
   */
  async createContact({ companyId, name, email }, { signal } = {}) {
    const content = await this.client.post(
      "/api/v1/employees", employeeWrite({ name, email, companyId }), { signal });
    const entity = Array.isArray(content) && content.length > 0 ? content[0] : content;
    const row = employeeRow(entity || {});
    if (!row.id) {
      throw new ApiError(
        "INTERNAL",
        "TANSS hat den Ansprechpartner ohne Kennung zurueckgegeben. Ob er angelegt "
        + "wurde, ist damit offen - bitte in TANSS nachsehen, bevor der Vorgang "
        + "wiederholt wird.",
      );
    }
    return { ...row, name: row.name || name, email: row.email || email };
  }

  /* ---------------------------------------------------------------- Mailablage */

  /**
   * Gelesene Nachricht an ein bestehendes Ticket haengen - und nachsehen, was daraus
   * wurde.
   *
   * Drei Stufen:
   *
   * 1. Hochladen mit GENAU EINEM `files`- und GENAU EINEM `descriptions`-Wert. Bei
   *    ungleicher Anzahl weist TANSS die Anfrage ab.
   * 2. Das `status` je Datei wird ausgewertet. Vom HTTP-Status wird NICHT auf den
   *    Ausgang der einzelnen Datei geschlossen - TANSS antwortet auch bei einer zu
   *    grossen Datei mit Erfolg.
   * 3. Gegenprobe: Steht die Nachricht danach in der Mailliste des Tickets, oder ist
   *    sie still als gewoehnliches Dokument gelandet?
   *
   * Der Mailbestand wird VOR dem Hochladen erhoben, damit die neue Nachricht ueber die
   * Differenz der Kennungen erkannt wird und nicht ueber den Betreff. Zwei Nachrichten
   * mit gleichem Betreff am selben Ticket waeren sonst nicht unterscheidbar.
   */
  async uploadEml(ticketId, { filename, blob, description }, { signal } = {}) {
    const before = await this._mailIds(ticketId, signal);
    const startedAt = Math.floor(Date.now() / 1000);

    const form = new FormData();
    form.append("files", blob, filename);
    form.append("descriptions", description);

    const content = await this.client.call(
      "POST", `/api/v1/tickets/${ticketId}/upload`, { body: form, signal });

    // Dieselbe Nachsicht wie in `createContact`: Die Antwort ist als Liste beschrieben,
    // in der Praxis kommt an dieser Schnittstelle auch ein einzelnes Objekt vor. Das ist
    // kein Zugestaendnis ins Blaue, sondern derselbe Fall an derselben Stelle.
    const results = Array.isArray(content)
      ? content.filter((item) => item)
      : (content && typeof content === "object" && "status" in content ? [content] : []);

    // Steht kein Ergebnis da, wird NICHT geraten, sondern nachgesehen. Genau diese Frage
    // ist offen - ob die Nachricht am Ticket haengt -, und genau dafuer gibt es die
    // Gegenprobe. Sie vorher zu ueberspringen war die falsche Reihenfolge: Sie meldete
    // einen Fehlschlag und bot die Wiederholung an, die eine angekommene Nachricht ein
    // zweites Mal angehaengt haette.
    if (results.length === 0) {
      return this._judgeByCheck(ticketId,
        { filename, blob, before, startedAt, shape: shapeOf(content), signal });
    }

    const result = results[0];
    const status = str(result.status).toUpperCase();
    if (status !== "OK") {
      return {
        status,
        fileName: str(result.name) || filename,
        sizeBytes: num(result.size) || blob.size,
        verified: false,
      };
    }

    const check = await this.verifyUpload(ticketId,
      { knownMailIds: before, since: startedAt, signal });
    return {
      status: "OK",
      confirmedBy: "response",
      fileName: str(result.name) || filename,
      sizeBytes: num(result.size) || blob.size,
      tanssMailId: check.tanssMailId,
      storedAsDocument: check.storedAsDocument,
      documentFileName: check.documentFileName,
      verified: check.checkedHistory || check.checkedDocuments,
    };
  }

  /**
   * Kein Ergebnis in der Antwort - also am Ticket nachsehen.
   *
   * Drei Ausgaenge, und sie sind wirklich verschieden:
   *
   *   1. Die Nachricht haengt am Ticket. Dann ist sie angekommen, und das wird gemeldet -
   *      mit dem Vermerk, woher die Bestaetigung stammt. Ein Erfolg aus der Gegenprobe
   *      ist ein Erfolg, aber er ist nicht derselbe wie eine Zusage des Servers, und die
   *      Oberflaeche sagt das.
   *   2. Die Gegenprobe lief und fand nichts. Dann ist der Fehlschlag BELEGT, nicht nur
   *      vermutet - und die Wiederholung ist gefahrlos.
   *   3. Die Gegenprobe kam selbst nicht durch. Dann ist nichts feststellbar, und genau
   *      das wird gesagt.
   *
   * Die Form der Antwort wird in allen Faellen genannt - nur ihre GESTALT, nie ihr
   * Inhalt. Beim naechsten Mal steht damit in der Meldung, was TANSS wirklich geschickt
   * hat, statt "ohne Ergebnisliste".
   */
  async _judgeByCheck(ticketId, { filename, blob, before, startedAt, shape, signal }) {
    const check = await this.verifyUpload(ticketId,
      { knownMailIds: before, since: startedAt, signal });

    if (check.tanssMailId !== null || check.storedAsDocument) {
      return {
        status: "OK",
        confirmedBy: "check",
        responseShape: shape,
        fileName: filename,
        sizeBytes: blob.size,
        tanssMailId: check.tanssMailId,
        storedAsDocument: check.storedAsDocument,
        documentFileName: check.documentFileName,
        verified: true,
      };
    }

    if (check.checkedHistory) {
      throw new ApiError(
        "INTERNAL",
        "TANSS hat den Upload bestaetigt, aber keine Ergebniszeile geliefert "
        + `(${shape}) - und am Ticket haengt die Nachricht nicht. Sie ist nicht `
        + "angekommen; ein zweiter Versuch ist gefahrlos.",
      );
    }

    throw new ApiError(
      "INTERNAL",
      "TANSS hat den Upload bestaetigt, aber keine Ergebniszeile geliefert "
      + `(${shape}). Die Gegenprobe am Ticket kam ebenfalls nicht durch, also laesst `
      + "sich nicht sagen, ob die Nachricht haengt - bitte in TANSS nachsehen, bevor "
      + "der Vorgang wiederholt wird.",
    );
  }

  /**
   * Prueft nach, was aus einer hochgeladenen Nachricht geworden ist.
   *
   * Beide Pruefungen sind bewusst NICHT toedlich: Scheitert eine, fehlt ein Nachweis,
   * aber die Nachricht kann trotzdem haengen. Das Ergebnis sagt ueber `checked*`, was
   * wirklich geprueft wurde.
   *
   * `knownMailIds === null` heisst "vorher nicht nachgesehen" und ist etwas anderes als
   * die leere Menge: Ohne Vergleichsstand gilt sonst JEDE am Ticket haengende Nachricht
   * als die soeben hochgeladene, und die Nachbearbeitung traefe die falsche.
   */
  async verifyUpload(ticketId, { knownMailIds, since, signal } = {}) {
    let storedAsDocument = false;
    let documentFileName = null;
    let checkedDocuments = false;
    try {
      const raw = (await this.client.get(
        `/api/v1/tickets/${ticketId}/documents`, { signal })) || [];
      checkedDocuments = true;
      for (const item of Array.isArray(raw) ? raw : []) {
        if (isFreshEmlDocument(item, ticketId, since)) {
          storedAsDocument = true;
          documentFileName = str(item.fileName);
          break;
        }
      }
    } catch {
      // Kein Negativnachweis. Das ist kein Grund, den Vorgang scheitern zu lassen.
    }

    let tanssMailId = null;
    let checkedHistory = false;
    const after = await this._mailIds(ticketId, signal);
    if (after !== null && knownMailIds !== null && knownMailIds !== undefined) {
      checkedHistory = true;
      const fresh = [...after].filter((id) => !knownMailIds.has(id)).sort((a, b) => a - b);
      // Mehrere neue Nachrichten gleichzeitig heisst: Der Mailroboter hat dazwischen
      // geliefert. Dann wird keine Kennung behauptet, statt die falsche zu melden.
      if (fresh.length === 1) tanssMailId = fresh[0];
    }

    return { tanssMailId, storedAsDocument, documentFileName, checkedHistory, checkedDocuments };
  }

  /**
   * Kennungen der am Ticket haengenden Nachrichten - `null`, wenn nicht abrufbar.
   *
   * `null` und "leere Menge" sind zwei verschiedene Aussagen: Die eine heisst "nicht
   * nachgesehen", die andere "nachgesehen, es haengt keine". Wer sie gleichsetzt, meldet
   * jede Nachricht als neu.
   */
  async _mailIds(ticketId, signal) {
    try {
      const content = (await this.client.get(
        `/api/v1/tickets/history/${ticketId}`, { signal })) || {};
      const mails = Array.isArray(content.mails) ? content.mails : [];
      return new Set(mails.map((mail) => num(mail.id)).filter((id) => id > 0));
    } catch {
      return null;
    }
  }
  /* ---------------------------------------------------------------- Einsaetze */

  /**
   * Ein einzelner Einsatz - `null` AUSSCHLIESSLICH, wenn es ihn nicht gibt.
   *
   * Eine leere Antwort mit Erfolgsstatus ist KEIN "gibt es nicht": Sie bedeutet, dass
   * etwas anderes schiefgelaufen ist. Wuerde sie als `null` durchgereicht, boete die
   * Terminmaske "Einsatz anlegen" an, obwohl laengst einer existiert - und es entstuende
   * ein Duplikat neben dem echten.
   */
  async getSupport(supportId, { signal } = {}) {
    let answer;
    try {
      answer = await this.client.get(`/api/v1/supports/${supportId}`,
        { wantMeta: true, signal });
    } catch (error) {
      if (error instanceof ApiError && error.code === "CONFLICT") return null;
      throw error;
    }
    const content = answer.content;
    if (!content || !content.id) {
      throw new ApiError(
        "INTERNAL",
        `TANSS lieferte fuer Einsatz ${supportId} eine leere Antwort mit Erfolgsstatus. `
        + "Das ist kein Nachweis, dass es ihn nicht gibt - der Vorgang bricht hier ab, "
        + "statt den Einsatz fuer geloescht zu halten.",
      );
    }
    return { support: supportRow(content), meta: answer.meta || {} };
  }

  /**
   * Einsaetze EINES Mitarbeiters im Zeitfenster.
   *
   * Drei Zusicherungen, die hier und nicht beim Aufrufer liegen:
   *
   * - Der Filter nennt den Mitarbeiter. Ein falscher Feldname wird stillschweigend
   *   ignoriert und liefert ALLE Mitarbeiter - die Terminmaske zeigte dann fremde
   *   Termine zur Auswahl an.
   * - KEIN Filter auf die Planungsart. Sonst verschwindet ein zur Leistung gewandelter
   *   Termin aus der Antwort und saehe aus wie geloescht; die Maske boete "Einsatz
   *   anlegen" an und erzeugte ein Duplikat neben der gebuchten Leistung.
   * - Die Zusatzinfos werden angefordert, ihr Erfolg aber nicht vorausgesetzt.
   */
  async listSupports(employeeId, from, to, { signal } = {}) {
    const { content } = await this.client.put("/api/v1/supports/list", {
      timeframe: { from: toUnix(from), to: toUnix(to) },
      employees: [employeeId],
      fetchMetaInfos: true,
    }, { retry: true, wantMeta: true, signal });

    const items = (Array.isArray(content) ? content : []).map(supportRow);
    return {
      items,
      // Kam wenigstens EIN Satz Zusatzinfos zurueck? Wenn nicht, beruht die Zuordnung
      // allein auf dem Zeitfenster - und die Maske sagt das, statt eine Kopplung zu
      // behaupten, die sie nicht geprueft hat.
      metaSeen: items.some((item) => Object.keys(item.metaInfos).length > 0),
    };
  }

  /**
   * Den TANSS-Einsatz zum geoeffneten Outlook-Termin finden.
   *
   * Die Stufen in dieser Reihenfolge, weil jede folgende schwaecher ist als die vorige:
   *
   * 1. **Hinterlegte Kennung**, gegengeprueft auf Mitarbeiter und Startzeit. Die
   *    Gegenprobe ist noetig, weil eine kopierte Termineinladung die Kennung des
   *    Originals mitbringt - ohne sie haenge die Maske am falschen Einsatz.
   * 2. **Kopplungswert** am Einsatz, verglichen nach der Regel des Servers.
   * 3. **Zeitfenster** ueber Beginn und Dauer. Eine Faustregel: Zwei gleich lange
   *    Termine desselben Technikers zur selben Minute sind nicht unterscheidbar.
   *
   * Mehrere Treffer ergeben NIE eine Entscheidung, immer eine Auswahl.
   */
  async resolveSupport(binding, { signal } = {}) {
    const windowMs = MATCH_WINDOW_MINUTES * 60_000;
    let stale = null;

    if (binding.supportId) {
      const found = await this.getSupport(binding.supportId, { signal });
      if (found && found.support.employeeId === binding.employeeId) {
        const startsWithin = !binding.start
          || Math.abs(found.support.date * 1000 - binding.start.getTime()) <= windowMs;
        if (startsWithin) {
          return this._match(found.support, found.meta, "property");
        }
        // Die Kennung zeigt auf einen Einsatz zu anderer Zeit - typisch fuer einen
        // kopierten Termin. Nicht uebernehmen, aber auch nicht verwerfen: als Kandidat
        // anbieten, falls sonst nichts gefunden wird.
        stale = this._candidate(found.support, new Map());
      }
    }

    // Stufe 2 und 3 spannen ihr Suchfenster um die Startzeit. Fehlt sie - ein noch
    // nicht gespeicherter Termin liefert keine -, gibt es kein Fenster und damit nichts
    // zu durchsuchen.
    if (!binding.start) {
      return {
        supportId: null, support: null, candidates: stale ? [stale] : [],
        source: "none", readOnly: false, metaInconclusive: false,
      };
    }

    const from = new Date(binding.start.getTime() - SUPPORT_SPAN_MS);
    const to = new Date(binding.start.getTime() + SUPPORT_SPAN_MS);
    const page = await this.listSupports(binding.employeeId, from, to, { signal });
    const metaInconclusive = !page.metaSeen;

    const uid = canonicalUid(binding.uid);
    if (uid) {
      const hits = page.items.filter((item) => syncGroupMatches(item.syncGroup, uid));
      if (hits.length === 1) return this._match(hits[0], {}, "syncgroup", metaInconclusive);
      if (hits.length > 1) return this._choice(hits, metaInconclusive);
    }

    const timed = page.items.filter((item) => this._matchesWindow(item, binding, windowMs));
    if (timed.length === 1) return this._match(timed[0], {}, "timematch", metaInconclusive);
    if (timed.length > 1) return this._choice(timed, metaInconclusive);

    return {
      supportId: null, support: null, candidates: stale ? [stale] : [],
      source: "none", readOnly: false, metaInconclusive,
    };
  }

  /**
   * Beginn UND Dauer im Fenster.
   *
   * Die Dauer trennt zwei Termine, die zur selben Minute beginnen - ohne sie waere ein
   * halbstuendiger Termin von einem ganztaegigen nicht zu unterscheiden.
   */
  _matchesWindow(support, binding, windowMs) {
    if (Math.abs(support.date * 1000 - binding.start.getTime()) > windowMs) return false;
    if (!binding.durationMinutes) return true;
    return Math.abs(support.duration - binding.durationMinutes) <= MATCH_WINDOW_MINUTES;
  }

  _match(support, meta, source, metaInconclusive = false) {
    const companies = linkedNames(meta, "companies");
    return {
      supportId: support.id,
      support,
      companyName: companies.get(support.companyId) || "",
      candidates: [],
      source,
      readOnly: !isEditableAppointment(support),
      metaInconclusive,
    };
  }

  _choice(hits, metaInconclusive) {
    return {
      supportId: null,
      support: null,
      candidates: hits.map((item) => this._candidate(item, new Map())),
      source: "none",
      readOnly: false,
      metaInconclusive,
    };
  }

  _candidate(support, companies) {
    const start = fromUnix(support.date);
    return {
      id: support.id,
      date: start ? start.toISOString() : null,
      duration: support.duration,
      companyId: support.companyId,
      companyName: companies.get(support.companyId) || "",
      ticketId: support.ticketId,
      planningType: support.planningType,
    };
  }

  /**
   * Die in der Maske gesetzten Felder auf den Einsatz zurueckschreiben.
   *
   * Vorher wird GELESEN, aus zwei Gruenden, die beide nicht verhandelbar sind: Ein zur
   * Leistung gewandelter Termin wird nicht geaendert, sondern gemeldet. Und der aktuelle
   * Satz der Zusatzinfos ist der Vergleichswert, ohne den die Obermengenregel wertlos
   * waere.
   *
   * Der Einsatz muss dem Handelnden gehoeren. Ohne diese Pruefung genuegte eine fremde
   * Einsatznummer, um Firma, Ticketbezug, Ort und Freitext im Kalendereintrag eines
   * Kollegen zu ueberschreiben - unbemerkt und nicht zuruecknehmbar.
   */
  async saveSupport(supportId, fields, { employeeId, signal } = {}) {
    const found = await this.getSupport(supportId, { signal });
    if (found === null) {
      throw new ApiError("CONFLICT", `Einsatz ${supportId} existiert nicht (mehr).`);
    }
    const current = found.support;
    if (current.employeeId && current.employeeId !== employeeId) {
      throw new ApiError(
        "APPOINTMENT_READONLY",
        "Dieser Einsatz gehoert einem anderen Mitarbeiter. Das Add-in aendert "
        + "ausschliesslich die eigenen Termine.",
      );
    }
    if (!isEditableAppointment(current)) {
      throw new ApiError("APPOINTMENT_READONLY");
    }

    const body = supportWrite(fields, { employeeId, forCreate: false });
    // Letzte Pruefung vor dem Absenden, gegen den SOEBEN gelesenen Satz.
    assertWriteSafe(body, { currentMeta: current.metaInfos });

    const content = await this.client.put(`/api/v1/supports/${supportId}`, body, { signal });
    if (!content || !content.id) {
      throw new ApiError(
        "INTERNAL",
        `TANSS hat die Aenderung an Einsatz ${supportId} mit Erfolg, aber ohne Inhalt `
        + "bestaetigt. Welche Felder uebernommen wurden, ist damit unbekannt - und weil "
        + "der Server nicht gesetzte Felder still verwirft, wird kein Erfolg gemeldet, "
        + "den wir nicht belegen koennen.",
      );
    }
    const saved = supportRow(content);
    const { applied, ignored } = appliedFields(body, content);
    return { supportId: saved.id || supportId, applied, ignored, syncGroup: saved.syncGroup };
  }

  /**
   * Einsatz zu einem noch nicht gekoppelten Termin anlegen.
   *
   * Nur in einer Installation OHNE Kalendersynchronisation. Laeuft ein solcher Dienst
   * gegen dieselbe Instanz, legt er den Einsatz an und das Add-in reichert ihn nur an -
   * beide anzulegen ergaebe zwei Eintraege fuer denselben Termin.
   *
   * Nur auf diesem Pfad gehen Zusatzinfos mit, und die Obermengenregel ist dabei
   * erfuellt, weil der aktuelle Satz eines noch nicht angelegten Einsatzes per
   * Definition leer ist.
   */
  async createSupport(fields, { employeeId, signal } = {}) {
    if (!(this.config.features && this.config.features.createSupport)) {
      throw new ApiError(
        "CONFLICT",
        "In dieser Installation legt die Kalendersynchronisation die Einsaetze an. Das "
        + "Add-in reichert sie nur an und legt selbst keine an.",
      );
    }

    const uid = canonicalUid(fields.uid);
    const body = supportWrite(fields, {
      employeeId,
      forCreate: true,
      syncGroup: uid ? formatSyncGroup(uid) : "",
    });
    assertWriteSafe(body, { currentMeta: {} });

    const content = await this.client.post("/api/v1/supports", body, { signal });
    if (!content || !content.id) {
      throw new ApiError(
        "INTERNAL",
        "TANSS hat den Einsatz ohne Kennung zurueckgegeben. Ob er angelegt wurde, ist "
        + "damit offen - bitte in TANSS nachsehen, bevor der Vorgang wiederholt wird.",
      );
    }
    const created = supportRow(content);
    const { applied, ignored } = appliedFields(body, content);
    return { supportId: created.id, applied, ignored, syncGroup: created.syncGroup };
  }
}

/**
 * Die Abteilungskennungen aus einer Mitarbeiterantwort - als Menge.
 *
 * Drei Schreibweisen, weil TANSS die Zuordnung an zwei Orten fuehrt und die Antwort je
 * nach Route anders geformt ist: ein einzelnes Feld am Mitarbeiter, eine Liste von
 * Kennungen, oder eine Liste von Zuordnungen mit eigenem Feld. Alle drei werden gelesen;
 * eine LEERE Menge heisst "nicht feststellbar" und wird von der Maske auch so behandelt.
 */
export function departmentIdsOf(payload) {
  const gefunden = new Set();
  const merken = (wert) => {
    const id = num(wert && typeof wert === "object" ? (wert.departmentId ?? wert.id) : wert);
    if (id) gefunden.add(id);
  };
  if (Array.isArray(payload.departmentAssignments)) payload.departmentAssignments.forEach(merken);
  if (Array.isArray(payload.departments)) payload.departments.forEach(merken);
  if (Array.isArray(payload.departmentIds)) payload.departmentIds.forEach(merken);
  merken(payload.departmentId);
  return [...gefunden];
}

/**
 * Die GESTALT einer Antwort in einem kurzen Satz - nie ihr Inhalt.
 *
 * Sie wandert in eine Fehlermeldung, die der Techniker liest und weitergibt. Deshalb nur
 * Typ, Laenge und Feldnamen: Feldnamen beschreiben die Schnittstelle, Werte beschrieben
 * den Vorgang - und der geht niemanden ausser dem Techniker etwas an.
 *
 * Ohne diese Angabe stand in der Meldung "ohne Ergebnisliste", und niemand konnte sagen,
 * was TANSS statt dessen geschickt hatte. Beim naechsten Mal steht es da.
 */
export function shapeOf(content) {
  if (content === null || content === undefined) return "leere Antwort";
  if (Array.isArray(content)) return `leere Liste (${content.length} Einträge)`;
  if (typeof content !== "object") return `${typeof content} statt Liste`;
  const felder = Object.keys(content).slice(0, 8);
  if (felder.length === 0) return "leeres Objekt statt Liste";
  return `Objekt statt Liste, Felder: ${felder.join(", ")}`;
}

/**
 * Ist dieses Dokument die gerade hochgeladene, still fehlabgelegte Nachricht?
 *
 * Zwei Zeitquellen, weil beide unzuverlaessig sein koennen: der Zeitstempel des
 * Dokuments und die Epoche im Dateinamen, den TANSS selbst bildet. Eine der beiden im
 * Fenster genuegt.
 */
export function isFreshEmlDocument(document, ticketId, since) {
  const match = EML_DOCUMENT.exec(str(document.fileName));
  if (!match || Number.parseInt(match[1], 10) !== ticketId) return false;
  const lower = since - CLOCK_SLACK_SECONDS;
  const upper = since + DOCUMENT_WINDOW_SECONDS;
  for (const stamp of [num(document.date), Number.parseInt(match[2], 10)]) {
    if (stamp && stamp >= lower && stamp <= upper) return true;
  }
  return false;
}
