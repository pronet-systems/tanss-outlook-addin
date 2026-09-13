/**
 * Der einzige Ort im Taskpane mit Zugriff auf `Office.*`.
 *
 * Warum eine einzige Tuer: Die Office-Schnittstelle ist die instabilste Abhaengigkeit des
 * Projekts. Sie unterscheidet sich zwischen klassischem Outlook, neuem Outlook, OWA und Mac,
 * zwischen Lese- und Verfassenmodus und zwischen Anforderungspaketen. Jede Weiche, die in
 * einer Seite steht statt hier, muss beim naechsten Host ein zweites Mal gefunden werden.
 *
 * Zweiter Grund: Alles, was von hier kommt, ist FREMDER Text - Betreff, Absendername,
 * Dateiname eines Anhangs. Es geht ausschliesslich ueber `ui.el()` und damit ueber
 * `textContent` in die Oberflaeche.
 *
 * Alle Aufrufe sind Promises. Die Office-Rueckrufe liefern `AsyncResult` mit einem Status;
 * der Status wird VOR dem Wert ausgewertet - ein leerer Wert ist nur bei `Succeeded` ein
 * leeres Ergebnis, sonst ein Fehler, der sonst als leeres Formular durchginge.
 */

/** Fehler aus der Office-Schicht, mit einem Code aus `T.errors`. */
export class OfficeError extends Error {
  constructor(code, detail = "") {
    super(code);
    this.name = "OfficeError";
    this.code = code;
    this.detail = detail;
  }
}

const READY_TIMEOUT_MS = 20_000;

/** @type {Promise<{host: string, platform: string}>|null} */
let readyPromise = null;

/**
 * Wartet auf Office.
 *
 * Mit Zeitgrenze: Bleibt `Office.onReady` aus - das passiert, wenn office.js vom CDN nicht
 * geladen wurde oder die Seite ausserhalb eines Outlook-Hosts geoeffnet ist -, dann soll das
 * Pane eine lesbare Meldung zeigen und nicht ewig einen Ladekreis.
 */
export function onReady() {
  if (readyPromise) return readyPromise;
  readyPromise = new Promise((resolve, reject) => {
    const office = globalThis.Office;
    if (!office || typeof office.onReady !== "function") {
      reject(new OfficeError("CLIENT_OFFICE_FAILED", "Office.onReady fehlt"));
      return;
    }
    const timer = window.setTimeout(() => {
      reject(new OfficeError("CLIENT_OFFICE_FAILED", "Office.onReady ohne Antwort"));
    }, READY_TIMEOUT_MS);

    office.onReady((info) => {
      window.clearTimeout(timer);
      resolve({
        host: info && info.host ? String(info.host) : "",
        platform: info && info.platform ? String(info.platform) : "",
      });
    });
  });
  return readyPromise;
}

/**
 * Ob die Office-Bibliothek ueberhaupt geladen ist.
 *
 * Der Unterschied zwischen "office.js fehlt" und "office.js ist da und meldet sich nicht"
 * ist der Unterschied zwischen zwei voellig verschiedenen Suchen: das erste ist eine
 * falsche Seitenadresse oder eine blockierte Quelle, das zweite ein Fehler der Einbettung.
 * Wer beides in denselben Satz packt, schickt den Administrator in die falsche Richtung.
 */
export function hasOfficeApi() {
  const office = globalThis.Office;
  return Boolean(office && typeof office.onReady === "function");
}

/** Kurzform fuer `Office.context.mailbox`; `null`, solange oder falls es sie nicht gibt. */
function mailbox() {
  const office = globalThis.Office;
  return office && office.context && office.context.mailbox ? office.context.mailbox : null;
}

/** Das geoeffnete Element. `null` heisst: das Pane wurde ohne Element geoeffnet. */
function item() {
  const box = mailbox();
  return box && box.item ? box.item : null;
}

/** Ob ueberhaupt ein Postfach-Host vorliegt. Entscheidet ueber die Faehigkeitspruefung. */
export function hasMailbox() {
  return mailbox() !== null;
}

/**
 * Anforderungspaket-Abfrage.
 *
 * Die einzige belastbare Quelle dafuer, was der laufende Host kann. Dokumentationswissen
 * ueber Versionsnummern ist Erwartung, diese Abfrage ist Messung - deshalb haengt jede
 * Weiche an ihr und nicht an einer Build-Nummer.
 */
export function isSetSupported(name, version) {
  const office = globalThis.Office;
  if (!office || !office.context || !office.context.requirements) return false;
  try {
    return office.context.requirements.isSetSupported(name, version) === true;
  } catch {
    return false;
  }
}

/** Liste der tatsaechlich gemeldeten Mailbox-Pakete - Rohstoff fuer `#/diagnose`. */
export function supportedMailboxSets() {
  const candidates = ["1.5", "1.6", "1.7", "1.8", "1.9", "1.10", "1.11", "1.12", "1.13", "1.14"];
  return candidates.filter((version) => isSetSupported("Mailbox", version));
}

/** Host, Plattform und Fassung, so wie der laufende Client sie meldet. */
export function hostInfo() {
  const office = globalThis.Office;
  const box = mailbox();
  const diagnostics = box && box.diagnostics ? box.diagnostics : null;
  return {
    host: office && office.context && office.context.host ? String(office.context.host) : "",
    platform:
      office && office.context && office.context.platform ? String(office.context.platform) : "",
    hostName: diagnostics && diagnostics.hostName ? String(diagnostics.hostName) : "",
    hostVersion: diagnostics && diagnostics.hostVersion ? String(diagnostics.hostVersion) : "",
    owaView: diagnostics && diagnostics.OWAView ? String(diagnostics.OWAView) : "",
  };
}

/**
 * Profil des angemeldeten Postfachs.
 *
 * `emailAddress` ist AUSSCHLIESSLICH Vorbelegung von Formularen und Hinweis fuer die
 * Kontowahl bei MSAL. Als Nachweis waere sie clientseitig faelschbar und damit eine
 * vollstaendige Autorisierungsumgehung - der Dienst leitet den Handelnden allein aus dem
 * geprueften Entra-Token ab.
 */
export function getUserProfile() {
  const box = mailbox();
  if (!box || !box.userProfile) return null;
  const profile = box.userProfile;
  return {
    displayName: profile.displayName ? String(profile.displayName) : "",
    emailAddress: profile.emailAddress ? String(profile.emailAddress) : "",
    // accountType liegt in Mailbox 1.6; auf aelteren Hosts ist der Wert leer.
    accountType: profile.accountType ? String(profile.accountType) : "",
    timeZone: profile.timeZone ? String(profile.timeZone) : "",
  };
}

/**
 * Ob das Postfach auf einem lokalen Exchange liegt.
 *
 * Daran haengt der gesamte Mailpfad: `enterprise` heisst, dass Microsoft Graph nicht zur
 * Verfuegung steht und die Nachricht aus Bordmitteln zusammengesetzt werden muss.
 * Ist der Wert leer (Host ohne Mailbox 1.6), nehmen wir NICHT `enterprise` an - dann
 * scheitert der Graph-Weg mit einer lesbaren Meldung, statt still eine Rekonstruktion
 * abzuliefern, die niemand angefordert hat.
 */
export function isEnterpriseMailbox() {
  const profile = getUserProfile();
  return profile !== null && profile.accountType === "enterprise";
}

/** "message", "appointment" oder `null`. Grundlage der Routenwahl in `main.js`. */
export function getItemType() {
  const current = item();
  const office = globalThis.Office;
  if (!current || !office || !office.MailboxEnums) return null;
  const types = office.MailboxEnums.ItemType;
  if (current.itemType === types.Message) return "message";
  if (current.itemType === types.Appointment) return "appointment";
  return null;
}

/* ------------------------------------------------------- Aufruf-Umhuellung */

/**
 * Macht aus einem Office-Rueckruf ein Promise.
 *
 * Statusauswertung VOR Wertauswertung: dieselbe Regel wie im Backend gegenueber TANSS.
 * Wer zuerst auf `asyncResult.value` schaut, haelt ein `undefined` aus einem Fehlerfall
 * fuer "es gibt nichts" - und laesst den Benutzer ein leeres Formular abschicken.
 */
function callAsync(fn, code = "CLIENT_OFFICE_FAILED") {
  return new Promise((resolve, reject) => {
    const office = globalThis.Office;
    const done = (asyncResult) => {
      if (!asyncResult || asyncResult.status !== office.AsyncResultStatus.Succeeded) {
        const error = asyncResult && asyncResult.error ? asyncResult.error : null;
        reject(new OfficeError(code, error ? `${error.code}: ${error.name}` : "unbekannt"));
        return;
      }
      resolve(asyncResult.value);
    };
    try {
      fn(done);
    } catch (error) {
      reject(new OfficeError(code, String(error && error.message ? error.message : error)));
    }
  });
}

/** Liest ein Verfassen-Feld (`subject`, `location`, …), das im Lesemodus schlicht ein Text ist. */
function readField(field) {
  if (field === null || field === undefined) return Promise.resolve("");
  if (typeof field === "string") return Promise.resolve(field);
  if (typeof field.getAsync === "function") {
    return callAsync((done) => field.getAsync(done)).then((value) =>
      value === null || value === undefined ? "" : String(value),
    );
  }
  return Promise.resolve("");
}

/* ------------------------------------------------------------------- Nachricht */

/** Eine Adresse in der Form, die `mime.js` und die Ticketmaske brauchen. */
function toAddress(entry) {
  if (!entry) return null;
  return {
    name: entry.displayName ? String(entry.displayName) : "",
    address: entry.emailAddress ? String(entry.emailAddress) : "",
  };
}

function toAddressList(entries) {
  return Array.isArray(entries) ? entries.map(toAddress).filter((a) => a && a.address) : [];
}

/**
 * Kopfdaten der geoeffneten Nachricht.
 *
 * Alles daraus ist rein lesend; nichts davon ist ein Nachweis. `itemId` ist die
 * Office-Kennung und muss fuer Graph erst mit `convertToRestId` gewandelt werden.
 */
export function getMessageContext() {
  const current = item();
  if (!current) throw new OfficeError("CLIENT_ITEM_MISSING");
  return {
    itemId: current.itemId ? String(current.itemId) : "",
    subject: typeof current.subject === "string" ? current.subject : "",
    normalizedSubject:
      typeof current.normalizedSubject === "string" ? current.normalizedSubject : "",
    from: toAddress(current.from) || toAddress(current.sender),
    to: toAddressList(current.to),
    cc: toAddressList(current.cc),
    dateTimeCreated: current.dateTimeCreated instanceof Date ? current.dateTimeCreated : null,
    internetMessageId: current.internetMessageId ? String(current.internetMessageId) : "",
    conversationId: current.conversationId ? String(current.conversationId) : "",
  };
}

/** Nachrichtentext. `coercion` ist "html" oder "text". */
export function getBody(coercion = "html") {
  const current = item();
  if (!current || !current.body) throw new OfficeError("CLIENT_ITEM_MISSING");
  const office = globalThis.Office;
  const type =
    coercion === "text" ? office.CoercionType.Text : office.CoercionType.Html;
  return callAsync((done) => current.body.getAsync(type, done)).then((value) =>
    value === null || value === undefined ? "" : String(value),
  );
}

/**
 * Anhangsliste der geoeffneten Nachricht (Metadaten, ohne Inhalt).
 *
 * Im Lesemodus liegt sie als `item.attachments` bereits vor; ein zusaetzlicher Aufruf von
 * `getAttachmentsAsync` waere eine Runde mehr fuer dieselben Daten.
 */
export function getAttachments() {
  const current = item();
  if (!current) throw new OfficeError("CLIENT_ITEM_MISSING");
  const list = Array.isArray(current.attachments) ? current.attachments : [];
  return list.map((attachment) => ({
    id: String(attachment.id),
    name: attachment.name ? String(attachment.name) : "",
    contentType: attachment.contentType ? String(attachment.contentType) : "",
    size: Number(attachment.size) || 0,
    isInline: attachment.isInline === true,
    attachmentType: attachment.attachmentType ? String(attachment.attachmentType) : "",
    contentId: attachment.contentId ? String(attachment.contentId) : "",
  }));
}

/**
 * Inhalt eines Anhangs.
 *
 * Braucht Mailbox 1.8. Die Weiche darauf sitzt in `mime.js`; hier wird nur ausgefuehrt,
 * damit die Faehigkeitspruefung genau eine Stelle hat.
 *
 * @returns {Promise<{format: string, content: string}>} `format` ist eine der Office-Marken
 *          base64 | eml | iCalendar | url.
 */
export function getAttachmentContent(attachmentId) {
  const current = item();
  if (!current || typeof current.getAttachmentContentAsync !== "function") {
    throw new OfficeError("CLIENT_OFFICE_FAILED", "getAttachmentContentAsync fehlt");
  }
  return callAsync((done) => current.getAttachmentContentAsync(attachmentId, done)).then(
    (value) => ({
      format: value && value.format ? String(value.format) : "",
      content: value && typeof value.content === "string" ? value.content : "",
    }),
  );
}

/* --------------------------------------------------------------------- Termin */

/**
 * Kopfdaten des geoeffneten Termins (Organisator-Verfassenmodus).
 *
 * Teilnehmer werden EINMAL beim Oeffnen gelesen. Auf `RecipientsChanged` wird bewusst
 * verzichtet: es hoebe die Manifest-Untergrenze auf Mailbox 1.7, und ein Handler muesste
 * beim Routenwechsel zuverlaessig abgemeldet werden - ein Abmelde-Handle mit falsch
 * gebundenem `this` ist dabei ein leicht uebersehener Fehler.
 */
export async function getAppointmentContext() {
  const current = item();
  if (!current) throw new OfficeError("CLIENT_ITEM_MISSING");

  const [subject, location, start, end] = await Promise.all([
    readField(current.subject),
    readField(current.location),
    readDateField(current.start),
    readDateField(current.end),
  ]);

  const durationMinutes =
    start instanceof Date && end instanceof Date
      ? Math.max(0, Math.round((end.getTime() - start.getTime()) / 60_000))
      : 0;

  return {
    subject,
    location,
    start,
    end,
    durationMinutes,
    organizer: toAddress(current.organizer),
  };
}

/** Datumsfeld, das im Lesemodus ein `Date` und im Verfassenmodus ein Async-Feld ist. */
function readDateField(field) {
  if (field instanceof Date) return Promise.resolve(field);
  if (field && typeof field.getAsync === "function") {
    return callAsync((done) => field.getAsync(done)).then((value) =>
      value instanceof Date ? value : value ? new Date(value) : null,
    );
  }
  return Promise.resolve(null);
}

/**
 * Speichert den geoeffneten Termin und liefert seine Kennung.
 *
 * Die Reihenfolge ist zwingend: ERST speichern, DANN die Kennung holen. Ein nie
 * gespeicherter Termin hat keine Kennung; `getItemIdAsync` antwortet dann mit
 * `ItemNotSaved`, und ohne diese Kennung gibt es keinen stabilen Schluessel, unter dem die
 * Kopplung zum TANSS-Einsatz abgelegt werden koennte.
 */
export async function saveAppointment() {
  const current = item();
  if (!current || typeof current.saveAsync !== "function") {
    throw new OfficeError("APPOINTMENT_NOT_SAVED", "saveAsync fehlt");
  }
  const value = await callAsync((done) => current.saveAsync(done), "APPOINTMENT_NOT_SAVED");
  return value ? String(value) : "";
}

/**
 * Kennung des geoeffneten Elements.
 *
 * Im Lesemodus steht sie als `item.itemId` bereit. Im Verfassenmodus (Termin) braucht es
 * `getItemIdAsync` aus Mailbox 1.8, und der Termin muss vorher gespeichert sein.
 */
export async function getItemId({ saveFirst = false } = {}) {
  const current = item();
  if (!current) throw new OfficeError("CLIENT_ITEM_MISSING");

  if (typeof current.itemId === "string" && current.itemId !== "") return current.itemId;

  if (saveFirst) {
    const savedId = await saveAppointment();
    if (savedId) return savedId;
  }
  if (typeof current.getItemIdAsync !== "function") {
    throw new OfficeError("APPOINTMENT_NOT_SAVED", "getItemIdAsync fehlt");
  }
  const value = await callAsync((done) => current.getItemIdAsync(done), "APPOINTMENT_NOT_SAVED");
  if (!value) throw new OfficeError("APPOINTMENT_NOT_SAVED");
  return String(value);
}

/* --------------------------------------------------------- Eigene Eigenschaften */

/**
 * Laedt die Custom Properties des geoeffneten Elements.
 *
 * Sie sind ausdruecklich nur ein BESCHLEUNIGER, nie die Wahrheit: beim Kopieren in ein
 * anderes Postfach gehen sie verloren. Die massgebliche Kopplung steht in TANSS unter
 * `metaInfos.SYNC_GROUP` und ist daraus jederzeit rekonstruierbar.
 *
 * @returns {Promise<{get: (k: string) => unknown, set: (k: string, v: unknown) => void,
 *                    remove: (k: string) => void, save: () => Promise<void>}>}
 */
export async function loadCustomProperties() {
  const current = item();
  if (!current || typeof current.loadCustomPropertiesAsync !== "function") {
    throw new OfficeError("CLIENT_OFFICE_FAILED", "loadCustomPropertiesAsync fehlt");
  }
  const bag = await callAsync((done) => current.loadCustomPropertiesAsync(done));
  return {
    get: (key) => bag.get(key),
    set: (key, value) => bag.set(key, value),
    remove: (key) => bag.remove(key),
    // saveAsync schreibt die gesamte Ablage; ohne diesen Aufruf ist jedes `set` wirkungslos.
    save: () => callAsync((done) => bag.saveAsync(done)).then(() => undefined),
  };
}

/* ------------------------------------------------------------------ Sonstiges */

/**
 * Wandelt eine Office-Kennung in eine REST-Kennung.
 *
 * Microsoft Graph versteht die Office-Kennung nicht. Die Wandlung ist rein lokal - sie
 * kostet keinen Netzaufruf und braucht keine Berechtigung.
 */
export function convertToRestId(itemId) {
  const box = mailbox();
  const office = globalThis.Office;
  if (!box || typeof box.convertToRestId !== "function") {
    throw new OfficeError("CLIENT_OFFICE_FAILED", "convertToRestId fehlt");
  }
  return String(box.convertToRestId(itemId, office.MailboxEnums.RestVersion.v2_0));
}

/**
 * Oeffnet eine Adresse im Browser des Benutzers.
 *
 * Nur `https:`. Eine aus Ticketnummer und Titel unkodiert zusammengesetzte Adresse darf
 * hier gar nicht erst geoeffnet werden koennen - die Pruefung steht deshalb an der
 * Ausgangstuer und nicht nur an der Stelle, die die Adresse baut.
 * Ohne `openBrowserWindow` (sehr alte Hosts) bleibt `window.open`; ohne beides bliebe der
 * Benutzer ohne Rueckmeldung, deshalb der Rueckgabewert.
 */
export function openInBrowser(url) {
  let parsed;
  try {
    parsed = new URL(String(url), window.location.href);
  } catch {
    return false;
  }
  if (parsed.protocol !== "https:") return false;

  const office = globalThis.Office;
  const ui = office && office.context && office.context.ui ? office.context.ui : null;
  if (ui && typeof ui.openBrowserWindow === "function") {
    ui.openBrowserWindow(parsed.href);
    return true;
  }
  const opened = window.open(parsed.href, "_blank", "noopener,noreferrer");
  return opened !== null;
}

/** Schliesst das Taskpane. Vorhanden ab Mailbox 1.5; fehlt es, bleibt das Pane offen. */
export function closePane() {
  const office = globalThis.Office;
  const ui = office && office.context && office.context.ui ? office.context.ui : null;
  if (ui && typeof ui.closeContainer === "function") {
    ui.closeContainer();
    return true;
  }
  return false;
}
