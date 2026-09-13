/**
 * Was sich das Pane merkt - und was ausdruecklich nicht.
 *
 * Es gibt keinen Dienst und keine Datenbank mehr. Alles, was ueber das Schliessen des
 * Panes hinaus gelten soll, liegt im oertlichen Speicher DIESES Arbeitsplatzes. Daraus
 * folgen drei Dinge, die man beim Lesen der Oberflaeche kennen muss:
 *
 * - **Es ist je Arbeitsplatz, nicht je Mitarbeiter.** Wer an zwei Rechnern arbeitet,
 *   hat zwei Vorliebenstaende. Das ist verschmerzbar bei Vorlieben und bei den zuletzt
 *   benutzten Tickets.
 * - **Die Dublettenwarnung wird dadurch schwaecher.** Sie verhindert zuverlaessig, dass
 *   MAN SELBST dieselbe Nachricht zweimal anhaengt. Dass ein Kollege sie bereits
 *   angehaengt hat, kann sie nicht wissen. Die Maske sagt das auch so - sie behauptet
 *   keine Vollstaendigkeit, die sie nicht hat.
 * - **Die Terminkopplung liegt NICHT hier, sondern am Termin.** Die Eigenschaft am
 *   Kalendereintrag wandert mit dem Postfach und ist an jedem Geraet dieselbe; dieser
 *   Speicher ist dafuer nur ein Zwischenspeicher, der jederzeit verloren gehen darf.
 *
 * Was hier NIE liegt: Betreffzeilen, Mailtexte, Anhaenge, Firmennamen, Ticketinhalte.
 * Von einer angehaengten Nachricht bleibt allein ein Streuwert der Nachrichtenkennung -
 * genug fuer "schon einmal angehaengt", zu wenig, um daraus einen Korrespondenzpartner
 * zu lesen.
 *
 * Jeder Zugriff faengt einen gesperrten Speicher ab. Ein privates Fenster, geleerte
 * Seitendaten oder ein Host, der den Speicher sperrt, machen das Pane langsamer im
 * Bedienen - nie funktionsunfaehig.
 */

/** Wie viele zuletzt benutzte Tickets gefuehrt werden. */
const RECENT_LIMIT = 10;

/** Wie viele Anhaengevermerke behalten werden, bevor der aelteste weicht. */
const ATTACHED_LIMIT = 500;

/** Wie lange ein Anhaengevermerk gilt. Danach ist die Warnung nicht mehr hilfreich. */
const ATTACHED_TTL_MS = 180 * 86_400_000;

/**
 * Die Schluessel, die eine Vorliebe tragen darf.
 *
 * Bewusst geschlossen: Der Speicher ist beliebig beschreibbar, und ein offener
 * Schluesselraum waere ein kleiner, kostenloser Datenspeicher fuer alles, was jemand
 * spaeter darin ablegen moechte - einschliesslich Dingen, die hier nicht liegen sollen.
 */
const ALLOWED_PREF_KEYS = new Set([
  "ticketTypeId",
  "ticketStatusId",
  "assignedToEmployeeId",
  "assignedToDepartmentId",
  "searchScope",
]);

/** Laengengrenze eines Vorliebenwertes. Die Werte sind Kennungen und kurze Woerter. */
const PREF_VALUE_MAX = 200;

/* ------------------------------------------------------------------ Zugriff */

/**
 * Streuwert einer Zeichenkette.
 *
 * Genug, um "dieselbe Nachricht" wiederzuerkennen, und zu wenig, um die Kennung daraus
 * zurueckzugewinnen. Bewusst ohne `crypto.subtle`: Das ist asynchron, und ein
 * asynchroner Aufruf in der Zeile, die eine Trefferliste zeichnet, waere eine
 * Fehlerquelle fuer einen Zweck, der keine kryptografische Staerke braucht.
 */
export function digest(text) {
  const value = String(text || "");
  let h1 = 0x811c9dc5;
  let h2 = 0x01000193;
  for (let i = 0; i < value.length; i += 1) {
    const c = value.charCodeAt(i);
    h1 = Math.imul(h1 ^ c, 0x01000193) >>> 0;
    h2 = Math.imul(h2 ^ ((c << 5) | (c >>> 3)), 0x85ebca6b) >>> 0;
  }
  return `${h1.toString(16).padStart(8, "0")}${h2.toString(16).padStart(8, "0")}`;
}

export class Store {
  /**
   * @param {{instance: string, employeeId: number}} scope
   *        Die Installation und der Mitarbeiter trennen die Bereiche voneinander. Ohne
   *        die Installation truege ein Termin aus einem Testsystem hier die Kopplung
   *        des Produktivsystems; ohne den Mitarbeiter saehe ein Techniker an einem
   *        geteilten Rechner die Vorbelegungen seines Kollegen.
   */
  constructor({ instance, employeeId }) {
    this.prefix = `tanss.addin.${instance || "unbekannt"}.${employeeId || 0}`;
  }

  _read(area, fallback) {
    try {
      const raw = globalThis.localStorage.getItem(`${this.prefix}.${area}`);
      if (!raw) return fallback;
      const parsed = JSON.parse(raw);
      return parsed === null || parsed === undefined ? fallback : parsed;
    } catch {
      return fallback;
    }
  }

  _write(area, value) {
    try {
      globalThis.localStorage.setItem(`${this.prefix}.${area}`, JSON.stringify(value));
      return true;
    } catch {
      // Gesperrt oder voll. Der Vorgang laeuft weiter; nur das Merken entfaellt.
      return false;
    }
  }

  /* ---------------------------------------------------------------- Vorlieben */

  /** Alle gemerkten Vorbelegungen. */
  prefs() {
    const stored = this._read("prefs", {});
    const result = {};
    for (const key of ALLOWED_PREF_KEYS) {
      const value = stored[key];
      if (typeof value === "string" && value !== "") result[key] = value;
    }
    return result;
  }

  /**
   * Merkt eine Vorbelegung. Unbekannte Schluessel werden verworfen, nicht angelegt.
   *
   * @returns {boolean} Ob der Wert uebernommen wurde.
   */
  setPref(key, value) {
    if (!ALLOWED_PREF_KEYS.has(key)) return false;
    const text = String(value === null || value === undefined ? "" : value)
      .slice(0, PREF_VALUE_MAX);
    const prefs = this._read("prefs", {});
    if (text === "") delete prefs[key];
    else prefs[key] = text;
    return this._write("prefs", prefs);
  }

  /** Eine gemerkte Kennung als Zahl - `null`, wenn sie fehlt oder Unsinn ist. */
  intPref(key) {
    const raw = (this.prefs()[key] || "").trim();
    return /^\d+$/.test(raw) ? Number.parseInt(raw, 10) || null : null;
  }

  /* ------------------------------------------------------- Zuletzt benutzt */

  /**
   * Die zuletzt benutzten Ticketnummern, neueste zuerst.
   *
   * Gespeichert wird NUR die Nummer. Die Titel holt die Maske bei Bedarf mit einem
   * einzigen Aufruf nach - so liegt kein Kundeninhalt im Ruhezustand, und ein
   * umbenanntes Ticket heisst sofort richtig.
   */
  recentTickets() {
    const list = this._read("recent", []);
    return Array.isArray(list)
      ? list.map((id) => Number(id)).filter((id) => Number.isInteger(id) && id > 0)
      : [];
  }

  /** Vermerkt ein Ticket als zuletzt benutzt. */
  touchTicket(ticketId) {
    const id = Number(ticketId);
    if (!Number.isInteger(id) || id <= 0) return false;
    const list = [id, ...this.recentTickets().filter((other) => other !== id)];
    return this._write("recent", list.slice(0, RECENT_LIMIT));
  }

  /* ------------------------------------------------------ Dublettenwarnung */

  /**
   * An welchen Tickets diese Nachricht von DIESEM Arbeitsplatz aus schon haengt.
   *
   * Die Warnung, die daraus entsteht, ist eine Warnung und keine Sperre: Es gibt
   * berechtigte Gruende, dieselbe Nachricht erneut abzulegen.
   */
  attachedTicketIds(internetMessageId) {
    if (!internetMessageId) return new Set();
    const key = digest(internetMessageId);
    const records = this._read("attached", {});
    const entry = records[key];
    if (!entry || !Array.isArray(entry.tickets)) return new Set();
    return new Set(entry.tickets.map((id) => Number(id)).filter((id) => id > 0));
  }

  /** Vermerkt, dass eine Nachricht an einem Ticket haengt. */
  rememberAttached(internetMessageId, ticketId) {
    if (!internetMessageId || !ticketId) return false;
    const key = digest(internetMessageId);
    const records = this._read("attached", {});
    const entry = records[key] || { tickets: [], at: 0 };
    if (!entry.tickets.includes(ticketId)) entry.tickets.push(ticketId);
    entry.at = Date.now();
    records[key] = entry;
    return this._write("attached", prune(records));
  }

  /* ------------------------------------------------- Terminkopplung (Cache) */

  /**
   * Die gemerkte Einsatznummer zu einem Termin - ein ZWISCHENSPEICHER.
   *
   * Die massgebliche Kopplung steht an zwei belastbareren Stellen: in der Eigenschaft
   * am Kalendereintrag und in TANSS selbst. Geht dieser Eintrag verloren, kostet das
   * eine Suchanfrage, keine Zuordnung.
   */
  appointmentLink(itemKey) {
    if (!itemKey) return null;
    const links = this._read("links", {});
    const entry = links[digest(itemKey)];
    return entry && Number(entry.supportId) > 0 ? entry : null;
  }

  /** Merkt die Kopplung eines Termins. */
  rememberAppointmentLink(itemKey, { supportId, uid = "", startUnix = 0, source = "" }) {
    if (!itemKey || !supportId) return false;
    const links = this._read("links", {});
    links[digest(itemKey)] = {
      supportId: Number(supportId),
      uid: String(uid || ""),
      startUnix: Number(startUnix) || 0,
      source: String(source || ""),
      at: Date.now(),
    };
    return this._write("links", prune(links));
  }

  /** Entfernt die gemerkte Kopplung eines Termins. */
  forgetAppointmentLink(itemKey) {
    if (!itemKey) return false;
    const links = this._read("links", {});
    delete links[digest(itemKey)];
    return this._write("links", links);
  }

  /** Wirft alles weg, was dieser Arbeitsplatz gemerkt hat. */
  clear() {
    for (const area of ["prefs", "recent", "attached", "links"]) {
      try {
        globalThis.localStorage.removeItem(`${this.prefix}.${area}`);
      } catch {
        // Nichts zu raeumen.
      }
    }
  }
}

/**
 * Haelt einen Bestand klein: Altes faellt weg, und ueber der Grenze weicht das
 * aelteste.
 *
 * Ohne diese Grenze waechst der Speicher unbegrenzt, und irgendwann schlaegt das
 * Schreiben fehl - nicht an der Stelle, die zu viel abgelegt hat, sondern an der
 * naechsten, die etwas merken will.
 */
function prune(records) {
  const cutoff = Date.now() - ATTACHED_TTL_MS;
  const entries = Object.entries(records)
    .filter(([, entry]) => Number(entry && entry.at) >= cutoff)
    .sort((a, b) => Number(b[1].at) - Number(a[1].at))
    .slice(0, ATTACHED_LIMIT);
  return Object.fromEntries(entries);
}
