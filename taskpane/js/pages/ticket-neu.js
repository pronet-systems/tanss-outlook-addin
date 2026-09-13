/**
 * Seite `#/ticket/neu`: aus der geoeffneten E-Mail ein TANSS-Ticket anlegen UND die Mail
 * im selben Vorgang daran ablegen.
 *
 * Warum genau EIN Knopf: Waeren Anlegen und Anhaengen zwei Client-Aufrufe und scheiterte
 * der zweite, existierte das Ticket bereits - der Benutzer bliebe aber auf dem Formular,
 * und sein naechster Klick erzeugte ein zweites Ticket. Hier gibt es einen Aufruf
 * (`POST /api/tickets`), der beides im Dienst erledigt, und genau einen Idempotenzschluessel
 * je Vorgang. Der Schluessel wird verworfen, sobald der Benutzer ein Feld aendert: derselbe
 * Schluessel mit anderem Inhalt beantwortet der Dienst mit CONFLICT, nie mit der alten
 * Antwort.
 *
 * Arbeitsteilung: Firma, Melder und die Auswahllisten loest der DIENST aus der
 * Absenderadresse auf; Betreff und Mailtext kommen parallel aus Office.js und bleiben im
 * Taskpane, bis tatsaechlich ein Ticket entsteht. Der Mailinhalt verlaesst den Browser
 * also fruehestens beim Absenden - und dann direkt nach TANSS, ohne Zwischenstation.
 *
 * Der zitierte Vorgaenger wird abgetrennt und eingeklappt angeboten, nicht weggeworfen:
 * sonst steht der komplette Verlauf in jedem Ticket desselben Threads noch einmal.
 */

import { T, errorText, formatBytes, t } from "../../i18n/de.js";
import * as api from "../api.js";
import * as mime from "../mime.js";
import * as office from "../office.js";
import * as ui from "../ui.js";

/** TANSS schneidet laengere Titel wortlos ab - deshalb kuerzen WIR, sichtbar und am Wort. */
const TITLE_MAX = 100;

/**
 * Schluessel der serverseitig gemerkten Vorbelegung (`POST /api/prefs`).
 *
 * Geschlossene Liste: der Dienst weist alles zurueck, was nicht darin steht. Bewusst
 * serverseitig statt im localStorage - der ist in Outlook-Taskpanes je nach Host
 * partitioniert oder gesperrt und roamt nicht zwischen Geraeten.
 */
const PREF_TYPE = "ticketTypeId";

/**
 * Die Stufen der Ticketprioritaet, wie TANSS sie zeigt: Ziffer und ebenso viele Sterne.
 *
 * Keine Worte. "hoch", "dringend" oder "niedrig" waeren eine Behauptung ueber eine
 * Richtung, die die Schnittstelle nicht hergibt - und eine verkehrt herum eingebaute
 * Skala setzte jedes dringende Ticket auf die falsche Stufe, ohne dass es auffiele.
 */
const PRIORITIES = Array.from({ length: 9 }, (unused, index) => {
  const stufe = index + 1;
  return { value: stufe, label: `${stufe} (${"*".repeat(stufe)})` };
});
const PREF_STATUS = "ticketStatusId";
const PREF_ASSIGNEE = "assignedToEmployeeId";
const PREF_DEPARTMENT = "assignedToDepartmentId";

/** Trennzeichen zwischen zwei Angaben einer Zeile. Satzzeichen, kein uebersetzbarer Text. */
const SEPARATOR = " · ";

/**
 * Ursachen, bei denen "Mail jetzt anhaengen" nichts aendert.
 *
 * Der Dienst legt die genaue Ursache als `mail.cause` neben den Sammelbegriff
 * `TICKET_CREATED_MAIL_FAILED`. Bei einer zu grossen oder unlesbaren Nachricht endete
 * jeder weitere Versuch genauso; ein Knopf, der nachweislich nichts bewirkt, ist
 * schlimmer als keiner.
 */
const RETRY_POINTLESS = new Set([
  "MAIL_TOO_LARGE",
  "MAIL_UNPARSABLE",
  "MAIL_ALREADY_ATTACHED",
  "MAIL_STORED_AS_DOCUMENT",
  "FORBIDDEN_NO_COMPANY_ACCESS",
]);

/* --------------------------------------------------------------------- Textarbeit */

/** Antwortpraefixe wie `AW: WG: Re[2]:` - dieselbe Regel wie `util/text.py`. */
const PREFIX_RE = /^(?:\s*(?:RE|AW|WG|FW|FWD|ANTW)\s*(?:\[\d+\])?\s*:)+\s*/i;

/**
 * Zeilen, ab denen zitiert wird. Jede muss am ZEILENANFANG stehen - sonst schneidet ein
 * "Von:" mitten im Fliesstext die halbe Mail weg.
 */
const TRAILER_RES = [
  /^-{2,}\s*Urspr(?:ü|ue)ngliche Nachricht\s*-{2,}$/i,
  /^-{2,}\s*Original Message\s*-{2,}$/i,
  /^-{2,}\s*Weitergeleitete Nachricht\s*-{2,}$/i,
  /^_{10,}$/,
  /^Am .{4,80} schrieb .{1,120}:$/,
  /^On .{4,80} wrote:$/,
];
const FROM_LINE_RE = /^(?:Von|From)\s*:\s*\S/i;
const FOLLOWUP_RE = /^(?:An|To|Gesendet|Sent|Datum|Date|Betreff|Subject)\s*:/i;

/**
 * Eine Zeile, hoechstens `limit` Zeichen, gekuerzt an der Wortgrenze.
 *
 * Das Ergebnis ist NIE laenger als `limit` - das Auslassungszeichen zaehlt mit. TANSS
 * schneidet laengere Titel wortlos ab, und zwar mitten im Wort.
 */
function shortenTitle(text, limit = TITLE_MAX) {
  const clean = String(text || "").replace(/\s+/g, " ").trim();
  if (clean.length <= limit) return clean;
  const cut = clean.slice(0, limit - 1);
  const space = cut.lastIndexOf(" ");
  return `${(space > limit / 2 ? cut.slice(0, space) : cut).trimEnd()}…`;
}

/**
 * Betreff zu Titel: Antwortpraefixe weg, dann kuerzen.
 *
 * Nur auf den BETREFF angewandt, nie auf den vom Benutzer getippten Titel - wer dort
 * bewusst "RE: ..." schreibt, soll es behalten duerfen.
 */
function titleFromSubject(subject, limit = TITLE_MAX) {
  const clean = String(subject || "").replace(/\s+/g, " ").trim().replace(PREFIX_RE, "");
  return shortenTitle(clean, limit);
}

/**
 * Trennt den eigenen Text vom zitierten Vorgaenger.
 *
 * Der Trailer wird zurueckgegeben statt verworfen: die Maske zeigt ihn eingeklappt, damit
 * niemand raten muss, was abgeschnitten wurde.
 */
function splitQuoteTrailer(text) {
  const lines = String(text || "").replace(/\r\n|\r/g, "\n").split("\n");
  const cut = findTrailerStart(lines);
  if (cut === null) return { body: lines.join("\n").trim(), trailer: "" };
  return {
    body: lines.slice(0, cut).join("\n").trim(),
    trailer: lines.slice(cut).join("\n").trim(),
  };
}

function findTrailerStart(lines) {
  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index].trim();
    if (line === "") continue;
    if (TRAILER_RES.some((pattern) => pattern.test(line))) return index;
    if (FROM_LINE_RE.test(line) && hasFollowupHeader(lines, index)) return index;
  }
  return quotedBlockStart(lines);
}

/**
 * Auf "Von:" muss binnen vier Zeilen "An:"/"Gesendet:" folgen.
 *
 * Ohne diese Bedingung schnitte eine Mail, die mit "Von: der Baustelle aus ..." beginnt,
 * ihren gesamten Inhalt weg.
 */
function hasFollowupHeader(lines, index) {
  for (let i = index + 1; i < Math.min(index + 5, lines.length); i += 1) {
    if (FOLLOWUP_RE.test(lines[i].trim())) return true;
  }
  return false;
}

/** Anfang des zusammenhaengenden `>`-Blocks am Ende der Mail. */
function quotedBlockStart(lines) {
  let start = null;
  for (let index = lines.length - 1; index >= 0; index -= 1) {
    const line = lines[index].trim();
    if (line.startsWith(">")) {
      start = index;
      continue;
    }
    if (line === "") continue;
    break;
  }
  return start;
}

/* ------------------------------------------------------------------- Hilfsmittel */

/** Ganzzahl oder `null`. Leere Auswahl, "0" und Unsinn ergeben alle `null`. */
function numberOrNull(value) {
  if (value === null || value === undefined || value === "") return null;
  const number = Number(value);
  return Number.isFinite(number) && number > 0 ? number : null;
}

/** Ob eine Auswahlliste den gemerkten Wert ueberhaupt noch enthaelt. */
function contains(options, id) {
  if (id === null || id === undefined) return false;
  return Array.isArray(options) && options.some((entry) => Number(entry.id) === Number(id));
}

/** `[{id,name,inactive}]` aus dem Dienst zu Optionen fuer `ui.select`. */
function toSelectOptions(entries) {
  return (Array.isArray(entries) ? entries : []).map((entry) => ({
    value: entry.id,
    label: entry.name || String(entry.id),
    disabled: entry.inactive === true,
  }));
}

/* ------------------------------------------------------------------------ Seite */

/**
 * @param {{root: HTMLElement, me: object, signal: AbortSignal, navigate: Function,
 *          showError: Function, clearError: Function, reload: Function}} ctx
 * @returns {Promise<() => void>} Aufraeumfunktion fuer den Routenwechsel.
 */
export async function render(ctx) {
  let message;
  try {
    message = office.getMessageContext();
  } catch (error) {
    ui.replace(ctx.root, ui.el("p", { class: "fatal-text", text: errorText(error) }));
    return () => {};
  }

  ui.replace(ctx.root, ui.loading(T.app.loading));

  // Beides parallel: der Dienst loest die Absenderadresse auf, Office liefert den Text.
  // Nacheinander waere es die Summe zweier Wartezeiten vor dem ersten sichtbaren Feld.
  const [contextResult, bodyText] = await Promise.all([
    api.post("api/mail/context", {
      json: {
        senderAddress: message.from ? message.from.address : "",
        senderName: message.from ? message.from.name : "",
        subject: message.subject,
        internetMessageId: message.internetMessageId,
      },
      signal: ctx.signal,
    }),
    office.getBody("text").catch(() => ""),
  ]);
  if (ctx.signal.aborted) return () => {};

  if (!contextResult.ok) {
    ctx.showError(contextResult.error);
    ui.replace(
      ctx.root,
      ui.el("div", { class: "stack" }, [
        ui.el("p", { class: "fatal-text", text: errorText(contextResult.error) }),
        ui.button({ label: T.app.retry, onClick: () => void ctx.reload() }),
      ]),
    );
    return () => {};
  }

  const data = contextResult.data || {};
  const prefs = ctx.me && ctx.me.prefs ? ctx.me.prefs : {};
  const quoted = splitQuoteTrailer(bodyText);

  const state = {
    company: data.company && data.company.id ? data.company : null,
    confidence: data.company ? String(data.company.confidence || "none") : "none",
    remitter: data.remitter || null,
    options: data.options || null,
    similar: Array.isArray(data.similarTickets) ? data.similarTickets : [],
    attachedTo: Array.isArray(data.alreadyAttachedTo) ? data.alreadyAttachedTo : [],
    title: data.title || titleFromSubject(message.subject),
    content: quoted.body,
    trailer: quoted.trailer,
    attachMail: true,
    internal: false,
    typeId: numberOrNull(prefs[PREF_TYPE]),
    // Nicht gemerkt: Die Prioritaet gehoert zum einzelnen Vorgang, nicht zur Gewohnheit.
    // Ein vor Wochen gewaehlter Wert stillschweigend wieder einzusetzen hiesse, jedem
    // Ticket eine Dringlichkeit mitzugeben, die niemand fuer dieses Ticket gewaehlt hat.
    priority: null,
    statusId: numberOrNull(prefs[PREF_STATUS]),
    assigneeId: numberOrNull(prefs[PREF_ASSIGNEE]),
    departmentId: numberOrNull(prefs[PREF_DEPARTMENT]),
    /** Leer heisst: der naechste Absendeversuch ist ein NEUER Vorgang. */
    idempotencyKey: "",
    submitting: false,
  };

  /* ---------------------------------------------------------------- Bausteine */

  const notices = ui.el("div", { class: "stack" });
  const companyBox = ui.el("div", { class: "stack" });
  const companySearchBox = ui.el("div", { class: "stack", props: { hidden: true } });
  const companyResults = ui.el("div", { class: "stack" });
  const remitterBox = ui.el("div", { class: "stack" });
  const remitterExtraBox = ui.el("div", { class: "stack", props: { hidden: true } });
  const remitterResults = ui.el("div", { class: "stack" });
  const optionsBox = ui.el("div", { class: "stack" });
  const actionBox = ui.el("div", { class: "stack" });
  const statusLine = ui.el("p", { class: "muted" });

  /** Jede Aenderung am Entwurf macht den Idempotenzschluessel ungueltig. */
  function markDirty() {
    state.idempotencyKey = "";
  }

  const titleInput = ui.input({
    value: state.title,
    onInput: () => {
      state.title = titleInput.value;
      markDirty();
    },
  });
  const contentInput = ui.textarea({
    value: state.content,
    rows: 8,
    onInput: () => {
      state.content = contentInput.value;
      markDirty();
    },
  });

  // Nur-lesendes Textfeld statt `<pre>`: es bricht um, und das Pane scrollt nie waagerecht.
  const trailerView = ui.el("textarea", {
    class: "textarea",
    props: { readOnly: true, rows: 8, value: state.trailer, hidden: true },
  });
  const trailerToggle = ui.button({
    label: T.ticketNeu.quotedShow,
    variant: "ghost",
    onClick: () => {
      trailerView.hidden = !trailerView.hidden;
      trailerToggle.textContent = trailerView.hidden
        ? T.ticketNeu.quotedShow
        : T.ticketNeu.quotedHide;
    },
  });

  const attachCheck = ui.checkbox({
    label: T.ticketNeu.attachMail,
    checked: state.attachMail,
    onChange: (event) => {
      state.attachMail = event.target.checked === true;
      markDirty();
    },
  });
  const internalCheck = ui.checkbox({
    label: T.ticketNeu.internal,
    checked: state.internal,
    onChange: (event) => {
      state.internal = event.target.checked === true;
      markDirty();
    },
  });

  /* ----------------------------------------------------------------- Hinweise */

  function confidenceHint() {
    if (!state.company) return "";
    if (state.confidence === "address") return T.ticketNeu.companyFromAddress;
    if (state.confidence === "domain") return T.ticketNeu.companyFromDomain;
    return "";
  }

  function renderNotices() {
    const children = [];
    if (state.attachedTo.length > 0) {
      children.push(
        ui.banner({
          tone: "warn",
          label: T.ticketAnhaengen.alreadyAttached,
          actionLabel: T.ticketNeu.similarSwitch,
          onAction: () => ctx.navigate("#/ticket/anhaengen"),
        }),
      );
    }
    // Warum etwas ausgefallen ist, steht jetzt DABEI. Frueher endete der Grund in einem
    // leeren `catch`, und die Meldung sagte nur, DASS eine Liste fehlt - der
    // Administrator fing bei null an, obwohl der Grund im Augenblick des Scheiterns
    // vorlag. "HTTP 404" und "HTTP 403" fuehren zu ganz verschiedenen Suchen.
    const ausfaelle = (state.options && state.options.failures) || [];
    if (ausfaelle.length > 0) {
      children.push(ui.banner({
        tone: "warn",
        label: `${T.ticketNeu.optionsFailures} ${ausfaelle
          .map((eintrag) => `${eintrag.path} — ${eintrag.reason}`)
          .join(" · ")}`,
      }));
    }
    if (state.similar.length > 0) {
      children.push(
        ui.banner({
          tone: "info",
          label: T.ticketNeu.similarHint,
          actionLabel: T.ticketNeu.similarSwitch,
          onAction: () => ctx.navigate("#/ticket/anhaengen"),
        }),
        ui.section({
          heading: T.ticketNeu.similarHeading,
          children: [
            ui.list({
              items: state.similar,
              renderItem: (ticket) =>
                ui.listRow({
                  title: `${ticket.id}${SEPARATOR}${ticket.title || ""}`,
                  subtitle: ticket.companyName || "",
                  badgeLabel: ticket.statusName || "",
                  onClick: () => ctx.navigate("#/ticket/anhaengen"),
                }),
            }),
          ],
        }),
      );
    }
    const warnings = ctx.me && Array.isArray(ctx.me.warnings) ? ctx.me.warnings : [];
    for (const warning of warnings) {
      children.push(ui.banner({ tone: "warn", label: warning }));
    }
    ui.replace(notices, ...children);
  }

  /* -------------------------------------------------------------------- Firma */

  const companySearch = ui.searchField({
    placeholder: T.app.searchPlaceholder,
    minChars: 2,
    onSearch: (term) => void searchCompanies(term),
  });
  ui.mount(companySearchBox, companySearch.root, companyResults);

  function renderCompany() {
    const hint = confidenceHint();
    ui.replace(
      companyBox,
      ui.el("div", { class: "field" }, [
        ui.el("div", { class: "field-label", text: T.ticketNeu.company }),
        ui.row([
          ui.chip({
            label: state.company ? state.company.name : T.ticketNeu.companyUnknown,
            title: hint,
            tone: state.company ? "strong" : "muted",
          }),
          ui.button({
            label: T.ticketNeu.companyChange,
            variant: "ghost",
            onClick: () => {
              companySearchBox.hidden = !companySearchBox.hidden;
              if (!companySearchBox.hidden) companySearch.focus();
            },
          }),
        ]),
        hint ? ui.el("div", { class: "field-hint", text: hint }) : null,
      ]),
      companySearchBox,
    );
  }

  async function searchCompanies(term) {
    if (term === "") {
      ui.clear(companyResults);
      return;
    }
    companySearch.setBusy(true);
    ui.replace(companyResults, ui.loading(T.app.searching));
    const result = await api.get("api/search/companies", {
      query: { q: term, limit: 25 },
      signal: ctx.signal,
    });
    if (ctx.signal.aborted) return;
    companySearch.setBusy(false);
    if (!result.ok) {
      ctx.showError(result.error);
      ui.clear(companyResults);
      return;
    }
    const items = Array.isArray(result.data.items) ? result.data.items : [];
    ui.replace(
      companyResults,
      result.data.tooMany ? ui.banner({ tone: "warn", label: T.app.tooMany }) : null,
      ui.list({
        items,
        renderItem: (company) =>
          ui.listRow({
            title: company.name || String(company.id),
            subtitle: [company.postCode, company.city].filter(Boolean).join(" "),
            onClick: () => pickCompany(company),
          }),
      }),
    );
  }

  /**
   * Eine andere Firma uebernehmen.
   *
   * Alles, was zur alten Firma gehoerte, muss mit ihr verschwinden - und zwar SICHTBAR,
   * nicht nur im Zustand. Frueher wurde `state.remitter` verworfen, die Trefferliste der
   * alten Firma aber stehen gelassen: Die Namen blieben anklickbar, und ein Klick hing
   * eine fremde Person an den neuen Kunden. Genau dieser Weg ist im ersten echten Lauf
   * beschritten worden.
   *
   * Dasselbe gilt fuer die aehnlichen Tickets: Sie stammen aus der Mailauswertung und
   * gehoeren der alten Firma. Stehen zu lassen, was nicht mehr gilt, ist schlimmer als
   * nichts anzuzeigen - es sieht wie eine Aussage ueber den neuen Kunden aus.
   */
  function pickCompany(company) {
    state.company = { id: company.id, name: company.name || String(company.id) };
    state.confidence = "none";
    state.remitter = null;
    state.similar = [];
    state.remitterUnverified = false;
    companySearchBox.hidden = true;
    companySearch.input.value = "";
    ui.clear(companyResults);
    clearRemitterSearch();
    markDirty();
    renderCompany();
    renderRemitter();
    renderNotices();
    renderActions();
    void loadOptions();
  }

  /** Die Meldersuche vollstaendig abraeumen - Eingabe, Treffer und aufgeklappter Bereich. */
  function clearRemitterSearch() {
    remitterExtraBox.hidden = true;
    remitterExtraBox.dataset.mode = "";
    remitterSearch.input.value = "";
    ui.clear(remitterResults);
  }

  /* -------------------------------------------------------------------- Melder */

  const remitterSearch = ui.searchField({
    placeholder: T.app.searchPlaceholder,
    minChars: 2,
    onSearch: (term) => void searchEmployees(term),
  });

  function renderRemitter() {
    const children = [];
    if (state.remitter) {
      children.push(
        ui.el("div", { class: "field" }, [
          ui.el("div", { class: "field-label", text: T.ticketNeu.remitter }),
          ui.row([
            ui.chip({ label: state.remitter.name || "", tone: "strong" }),
            ui.button({
              label: T.app.search,
              variant: "ghost",
              onClick: () => toggleRemitterExtra("search"),
            }),
          ]),
          state.remitter.email
            ? ui.el("div", { class: "field-hint", text: state.remitter.email })
            : null,
        ]),
      );
    } else {
      children.push(
        ui.el("div", { class: "field" }, [
          ui.el("div", { class: "field-label", text: T.ticketNeu.remitter }),
          ui.banner({ tone: "info", label: T.ticketNeu.remitterUnknown }),
          ui.row([
            ui.button({
              label: T.app.search,
              variant: "ghost",
              onClick: () => toggleRemitterExtra("search"),
            }),
          ]),
        ]),
      );
    }
    ui.replace(remitterBox, ...children, remitterExtraBox);
  }

  /** Schaltet zwischen Suchfeld und Anlegen-Formular um; ein zweiter Klick klappt zu. */
  function toggleRemitterExtra(mode) {
    if (remitterExtraBox.dataset.mode === mode && remitterExtraBox.hidden === false) {
      remitterExtraBox.hidden = true;
      return;
    }
    remitterExtraBox.dataset.mode = mode;
    remitterExtraBox.hidden = false;
    ui.replace(remitterExtraBox, remitterSearch.root, remitterResults);
    remitterSearch.focus();
  }

  async function searchEmployees(term) {
    if (term === "") {
      ui.clear(remitterResults);
      return;
    }

    // Ohne gewaehlte Firma wird gar nicht erst gesucht. Eine firmenlose Suche liefe ueber
    // ALLE Kunden, und jeder Treffer waere anklickbar - ein Melder, der nirgends
    // hingehoert. Erst die Firma, dann die Person.
    if (!state.company) {
      ui.replace(remitterResults, ui.banner({ tone: "info", label: T.ticketNeu.companyFirst }));
      return;
    }

    // Die Firma wird FESTGEHALTEN, nicht nachgeschlagen. Zwischen Absenden und Antwort
    // kann der Techniker die Firma gewechselt haben; eine Antwort zur alten Firma darf in
    // der neuen Maske weder erscheinen noch anklickbar sein. Ein `signal` faengt das
    // nicht: Es faellt beim Seitenwechsel, nicht beim Firmenwechsel.
    const gesuchteFirma = state.company.id;

    remitterSearch.setBusy(true);
    ui.replace(remitterResults, ui.loading(T.app.searching));
    const result = await api.get("api/search/employees", {
      query: { q: term, companyId: gesuchteFirma },
      signal: ctx.signal,
    });
    if (ctx.signal.aborted) return;
    if (!state.company || state.company.id !== gesuchteFirma) {
      ui.clear(remitterResults);
      return;
    }
    remitterSearch.setBusy(false);
    if (!result.ok) {
      ctx.showError(result.error);
      ui.clear(remitterResults);
      return;
    }
    const items = Array.isArray(result.data.items) ? result.data.items : [];

    // Die Trefferliste ist bereits im Repository nach der Firma gefiltert. Dass die
    // Zuordnung nicht in jedem Fall FESTSTELLBAR ist, wird gesagt statt verschwiegen:
    // Nennt TANSS an einer Zeile keine Firma, laesst sich nicht pruefen, ob sie passt.
    state.remitterUnverified = result.data.unverified === true;

    ui.replace(
      remitterResults,
      result.data.tooMany ? ui.banner({ tone: "warn", label: T.app.tooMany }) : null,
      state.remitterUnverified
        ? ui.banner({ tone: "warn", label: T.ticketNeu.remitterUnverified })
        : null,
      ui.list({
        items,
        renderItem: (employee) =>
          ui.listRow({
            title: employee.name || String(employee.id),
            subtitle: employee.email || "",
            onClick: () => {
              // Zweite Sperre, unmittelbar am Klick. Die erste steht im Repository; diese
              // hier faengt den Fall ab, dass eine Liste aus anderem Grund stehen blieb.
              if (!state.company || state.company.id !== gesuchteFirma) {
                clearRemitterSearch();
                return;
              }
              state.remitter = employee;
              remitterExtraBox.hidden = true;
              markDirty();
              renderRemitter();
            },
          }),
      }),
    );
  }

  /**
   * Unbekannten Absender als Ansprechpartner anlegen.
   *
   * Das ist die haeufigste Sackgasse des Arbeitsablaufs: TANSS verlangt bei vielen Firmen
   * einen Melder, und wer ihn nicht hat, muesste sonst die Maske verlassen, in TANSS
   * wechseln, den Kontakt anlegen und von vorn beginnen.
   */
  function clampSelections() {
    const options = state.options || {};
    if (!contains(options.types, state.typeId)) state.typeId = null;
    if (!contains(options.states, state.statusId)) state.statusId = null;
    if (!contains(options.technicians, state.assigneeId)) state.assigneeId = null;
    // Ein gemerkter Wert, der nicht mehr zur Auswahl steht, wird gestrichen. Bei der
    // Abteilung ist das der Regelfall: Sie haengt am Techniker, und mit ihm wechselt sie.
    if (!contains(options.departments, state.departmentId)) state.departmentId = null;

    const defaults = ctx.me && ctx.me.defaults ? ctx.me.defaults : {};
    const fallbackType = numberOrNull(defaults.ticketTypeId);
    if (state.typeId === null && contains(options.types, fallbackType)) state.typeId = fallbackType;
  }

  function renderOptions() {
    const options = state.options || {};
    // Ohne Typen kein Typfeld. Ein Auswahlfeld ohne Auswahl sieht nach einem Ausfall aus
    // und laedt zum Suchen ein, wo es nichts zu finden gibt - und TANSS nimmt ein Ticket
    // ohne Typ an. Wer Typen anbieten will, gibt sie im Manifest mit; die Konfiguration
    // taugt dafuer nur bei eigener Ablage.
    const typen = Array.isArray(options.types) ? options.types : [];
    const typeSelect = typen.length === 0 ? null : ui.select({
      options: toSelectOptions(typen),
      value: state.typeId,
      placeholder: T.app.none,
      onChange: () => {
        state.typeId = numberOrNull(typeSelect.value);
        markDirty();
        void loadOptions();
      },
    });
    // Ziffer und Sterne, kein einziges Wort - genau so zeigt TANSS es selbst. Ob 1 oder
    // 9 "hoch" bedeutet, geht aus der Schnittstelle NICHT hervor: Verglichen wird
    // nirgends groesser/kleiner, und die einzige Richtungsangabe der Dokumentation
    // gehoert dem Rueckruf, nicht dem Ticket. Wer keine Richtung behauptet, kann sie
    // nicht verkehrt herum behaupten; die Sterne bilden die Zahl ab, sie deuten sie nicht.
    const prioritySelect = ui.select({
      options: PRIORITIES,
      value: state.priority,
      placeholder: T.ticketNeu.priorityDefault,
      onChange: () => {
        state.priority = numberOrNull(prioritySelect.value);
        markDirty();
      },
    });
    const statusSelect = ui.select({
      options: toSelectOptions(options.states),
      value: state.statusId,
      placeholder: T.app.none,
      onChange: () => {
        state.statusId = numberOrNull(statusSelect.value);
        markDirty();
      },
    });
    const assigneeSelect = ui.select({
      options: toSelectOptions(options.technicians),
      value: state.assigneeId,
      placeholder: T.app.none,
      onChange: () => {
        state.assigneeId = numberOrNull(assigneeSelect.value);
        markDirty();
        void loadOptions();
      },
    });
    const departmentSelect = ui.select({
      options: toSelectOptions(options.departments),
      value: state.departmentId,
      placeholder: T.app.none,
      onChange: () => {
        state.departmentId = numberOrNull(departmentSelect.value);
        markDirty();
      },
    });

    ui.replace(
      optionsBox,
      typeSelect ? ui.field({ label: T.ticketNeu.type, control: typeSelect }) : null,
      ui.field({ label: T.ticketNeu.status, control: statusSelect }),
      ui.field({ label: T.ticketNeu.priority, control: prioritySelect }),
      ui.field({
        label: T.ticketNeu.assignee,
        control: assigneeSelect,
        required: options.forceAssignment === true,
      }),
      ui.field({ label: T.ticketNeu.department, control: departmentSelect }),
      internalCheck.root,
    );
  }

  /* ------------------------------------------------------------------ Absenden */

  /** Anzahl echter Anhaenge; eingebettete Bilder sind fuer den Benutzer keine. */
  function attachmentCount() {
    try {
      return office.getAttachments().filter((attachment) => !attachment.isInline).length;
    } catch {
      return 0;
    }
  }

  function attachmentLabel(count) {
    if (count === 0) return T.mail.attachmentsNone;
    if (count === 1) return T.mail.attachmentsOne;
    return t("mail.attachmentsMany", { n: count });
  }

  function renderActions() {
    ui.replace(
      actionBox,
      attachCheck.root,
      ui.el("p", { class: "muted", text: attachmentLabel(attachmentCount()) }),
      mime.currentSource() === "officejs"
        ? ui.el("p", { class: "muted", text: T.mail.reconstructed })
        : null,
      statusLine,
      ui.button({
        label: T.ticketNeu.submit,
        variant: "primary",
        disabled: state.company === null || state.submitting,
        title: state.company === null ? T.ticketNeu.companyUnknown : "",
        onClick: () => void submit(),
      }),
    );
  }

  async function submit() {
    if (state.submitting) return;
    ctx.clearError();

    const title = shortenTitle(state.title, TITLE_MAX);
    if (title === "") {
      ctx.showError({ code: "TITLE_REQUIRED" });
      return;
    }
    if (!state.company) return;
    if (state.options && state.options.remitterRequired === true && !state.remitter) {
      ctx.showError({ code: "REMITTER_REQUIRED" });
      return;
    }

    state.submitting = true;
    ui.setBusy(ctx.root, true);
    statusLine.textContent = T.ticketNeu.submitting;

    // `ctx.root` ist `app.view` und ueberlebt jeden Seitenwechsel. Bleibt die Sperre nach
    // einem Abbruch - Klick auf "Diagnose" im Kopf, Wechsel ueber das Menueband - stehen,
    // traegt der Knoten `is-busy` fuer immer, und das Pane nimmt auf KEINER Seite mehr
    // einen Klick an; der Techniker kann es dann nur noch schliessen.
    try {
      await sendTicket(title);
    } finally {
      if (state.submitting) finishSubmit();
    }
  }

  /** Der eigentliche Absendeweg. Die Sperre nimmt ausschliesslich `submit()` zurueck. */
  async function sendTicket(title) {
    const form = new FormData();
    // `linkTypeId`/`linkId` setzt der DIENST aus der companyId; das Tripel ist eine
    // TANSS-Eigenheit und gehoert nicht in eine Oberflaeche.
    form.set(
      "ticket",
      JSON.stringify({
        companyId: state.company.id,
        remitterId: state.remitter ? state.remitter.id : null,
        title,
        content: state.content,
        typeId: state.typeId,
        priority: state.priority,
        statusId: state.statusId,
        assignedToEmployeeId: state.assigneeId,
        assignedToDepartmentId: state.departmentId,
        internal: state.internal,
      }),
    );

    if (state.attachMail) {
      statusLine.textContent = T.mail.fetching;
      const mail = await mime.fetchMessageMime({ maxBytes: maxEmlBytes() });
      if (ctx.signal.aborted) return;
      if (!mail.ok) {
        // Kein stilles Anlegen ohne Mail: wer die Mail am Ticket haben wollte und sie nicht
        // bekommt, entscheidet selbst, ob das Ticket trotzdem entstehen soll.
        finishSubmit();
        ctx.showError(mail.error);
        return;
      }
      statusLine.textContent = t("mail.sizeInfo", { size: formatBytes(mail.data.bytes) });
      form.set("eml", mail.data.blob, mail.data.filename);
      form.set("internetMessageId", mail.data.internetMessageId || "");
      form.set("subject", mail.data.subject || "");
      form.set("source", mail.data.source);
      for (const warning of mail.data.warnings) {
        ui.mount(notices, ui.banner({ tone: "warn", label: warning }));
      }
    }

    if (state.idempotencyKey === "") state.idempotencyKey = api.newIdempotencyKey();
    statusLine.textContent = T.ticketNeu.submitting;

    const result = await api.post("api/tickets", {
      form,
      idempotencyKey: state.idempotencyKey,
      signal: ctx.signal,
    });
    if (ctx.signal.aborted) return;
    finishSubmit();

    if (!result.ok) {
      ctx.showError(result.error);
      return;
    }
    rememberPrefs();
    renderSuccess(result.data || {});
  }

  function finishSubmit() {
    state.submitting = false;
    ui.setBusy(ctx.root, false);
    statusLine.textContent = "";
    // `setBusy` gibt ALLE Bedienelemente wieder frei - auch die, die gesperrt bleiben
    // muessen. Deshalb wird der Handlungsbereich danach neu gezeichnet.
    renderActions();
  }

  function maxEmlBytes() {
    return ctx.me && ctx.me.limits ? Number(ctx.me.limits.maxEmlBytes) || 0 : 0;
  }

  /**
   * Merkt Typ, Status und Zuweisung fuer das naechste Mal.
   *
   * Bewusst OHNE das Abbruchsignal der Seite: der Aufruf laeuft nach dem erfolgreichen
   * Anlegen, und genau dann wechselt die Ansicht. Ein abgebrochenes Merken waere kein
   * Schaden, aber eben auch kein Nutzen.
   */
  function rememberPrefs() {
    const pairs = [
      [PREF_TYPE, state.typeId],
      [PREF_STATUS, state.statusId],
      [PREF_ASSIGNEE, state.assigneeId],
      [PREF_DEPARTMENT, state.departmentId],
    ];
    for (const [key, value] of pairs) {
      if (value === null) continue;
      void api.post("api/prefs", { json: { key, value: String(value) } });
    }
  }

  /* --------------------------------------------------------------- Erfolgsseite */

  /**
   * Ergebnis eines nicht atomaren Vorgangs.
   *
   * TANSS kennt keinen Aufruf, der Ticket und Mail zusammen schreibt. Scheitert nur die
   * Ablage, bleibt das Ticket bestehen - und die Seite sagt genau das und bietet den
   * zweiten Aufruf noch einmal an, statt den Benutzer raten zu lassen.
   */
  function renderSuccess(created) {
    const mail = created.mail || {};
    const children = [
      ui.el("h2", {
        class: "section-heading",
        text: t("ticketNeu.successHeading", { id: created.ticketId }),
      }),
    ];

    if (mail.status === "attached") {
      // Woher die Bestaetigung stammt, ist kein Detail: Bei einer Antwort ohne
      // Ergebniszeile hat die Gegenprobe am Ticket entschieden. Der Techniker soll das
      // wissen - und vor allem, dass er NICHT nachhelfen muss. Genau dort lag die Gefahr:
      // Frueher stand hier ein Fehlschlag mit dem Angebot, es noch einmal zu versuchen,
      // und eine angekommene Nachricht waere ein zweites Mal angehaengt worden.
      children.push(ui.banner({
        tone: "ok",
        label: mail.confirmedBy === "check"
          ? T.ticketNeu.successMailByCheck
          : T.ticketNeu.successMailAttached,
      }));
    } else if (mail.status === "skipped") {
      children.push(ui.banner({ tone: "info", label: T.ticketNeu.successMailSkipped }));
    } else {
      const helps = !RETRY_POINTLESS.has(String(mail.cause || ""));
      children.push(
        ui.banner({
          tone: "warn",
          label: mail.message || T.ticketNeu.successMailFailed,
          actionLabel: helps ? T.ticketNeu.retryMail : "",
          onAction: helps ? () => void retryMail(created.ticketId) : null,
        }),
      );
    }

    if (created.url) {
      children.push(
        ui.button({
          label: T.app.openInTanss,
          variant: "primary",
          onClick: () => {
            if (!office.openInBrowser(created.url)) ctx.showError({ code: "CLIENT_OFFICE_FAILED" });
          },
        }),
      );
    }
    children.push(
      ui.button({ label: T.app.closePane, variant: "ghost", onClick: () => office.closePane() }),
    );
    ui.replace(ctx.root, ui.el("div", { class: "stack" }, children));
  }

  /** Nur der zweite Aufruf - die Dublettentabelle des Dienstes haelt ihn sauber. */
  async function retryMail(ticketId) {
    ctx.clearError();
    const mail = await mime.fetchMessageMime({ maxBytes: maxEmlBytes() });
    if (ctx.signal.aborted) return;
    if (!mail.ok) {
      ctx.showError(mail.error);
      return;
    }
    const form = new FormData();
    form.set("eml", mail.data.blob, mail.data.filename);
    form.set("internetMessageId", mail.data.internetMessageId || "");
    form.set("subject", mail.data.subject || "");
    form.set("source", mail.data.source);
    form.set("internal", state.internal ? "true" : "false");

    const result = await api.post(`api/tickets/${ticketId}/mail`, {
      form,
      idempotencyKey: api.newIdempotencyKey(),
      signal: ctx.signal,
    });
    if (ctx.signal.aborted) return;
    if (!result.ok) {
      ctx.showError(result.error);
      return;
    }
    renderSuccess({ ticketId, url: "", mail: { status: "attached" } });
  }

  /* ------------------------------------------------------------------- Aufbau */

  ui.replace(
    ctx.root,
    notices,
    ui.section({ heading: T.ticketNeu.heading, children: [companyBox, remitterBox] }),
    // Titel und Text nehmen den meisten Platz ein und sind in aller Regel schon richtig
    // vorbelegt. Wer sie nicht aendert, soll nicht an ihnen vorbeiscrollen muessen.
    //
    // Die Kopfzeile zeigt den aktuellen Titel mit: Eine zugeklappte Ziehharmonika, die
    // verschweigt, was in ihr steht, waere schlechter als gar keine - man muesste sie
    // oeffnen, nur um zu sehen, ob man sie oeffnen wollte.
    ui.collapsible({
      summary: ui.el("span", {}, [
        ui.el("span", { text: T.ticketNeu.textFold }),
        ui.el("span", { class: "fold-sub", text: state.title || "" }),
      ]),
      children: [
        ui.field({ label: T.ticketNeu.title, control: titleInput, required: true }),
        ui.charCounter(titleInput, TITLE_MAX),
        ui.field({ label: T.ticketNeu.content, control: contentInput }),
        state.trailer ? ui.row([trailerToggle]) : null,
        state.trailer ? trailerView : null,
      ],
    }),
    optionsBox,
    actionBox,
  );

  renderNotices();
  renderCompany();
  renderRemitter();
  renderActions();

  if (state.options) {
    clampSelections();
    renderOptions();
  } else {
    await loadOptions();
  }

  // Aufraeumen: die entprellten Suchfelder haben Zeitgeber, die sonst nach dem
  // Routenwechsel in eine abgeraeumte Oberflaeche schreiben.
  return () => {
    companySearch.dispose();
    remitterSearch.dispose();
  };
}
