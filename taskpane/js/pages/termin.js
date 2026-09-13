/**
 * Seite `#/termin`: den TANSS-Einsatz zum geoeffneten Outlook-Termin finden und pflegen.
 *
 * Die Reihenfolge am Anfang ist nicht verhandelbar: ERST `saveAsync`, DANN `getItemIdAsync`
 * (beides ueber `office.getItemId({saveFirst:true})`). Ein nie gespeicherter Termin hat
 * keine Kennung; `getItemIdAsync` antwortet dann mit `ItemNotSaved`, und ohne diese Kennung
 * gibt es keinen stabilen Schluessel, unter dem die Kopplung abgelegt werden koennte - die
 * Seite haette dann nichts, woran sie den Einsatz wiedererkennt.
 *
 * Der sichtbare Terminkoerper wird NIE angefasst. Ein vorhandener `{TANSS:<base64>}`-Marker
 * des alten Add-ins wird vom Dienst GELESEN und zur Vorbelegung benutzt, aber nie
 * geschrieben und nie entfernt - sonst verloere die Kalendersynchronisation des
 * Schwesterprojekts ihre Zuordnung. Teilnehmer, also Kunden, sehen diesen Marker; genau
 * deshalb schreiben wir ihn nicht.
 *
 * Die Custom Properties `tanssSupportId`/`tanssInstance` sind ausdruecklich nur ein
 * BESCHLEUNIGER: sie gehen beim Kopieren in ein anderes Postfach verloren. Massgeblich ist
 * `metaInfos.SYNC_GROUP` in TANSS, und das liest der Dienst - er schreibt es im Regelbetrieb
 * nicht.
 */

import { T, errorText, t } from "../../i18n/de.js";
import * as api from "../api.js";
import * as office from "../office.js";
import * as ui from "../ui.js";

/** Namen der eigenen Eigenschaften am Termin. Beide zusammen, nie eine allein. */
const PROP_SUPPORT_ID = "tanssSupportId";
const PROP_INSTANCE = "tanssInstance";

/** Trennzeichen zwischen zwei Angaben einer Zeile. Satzzeichen, kein uebersetzbarer Text. */
const SEPARATOR = " · ";

/** Ganzzahl oder `null`. Leere Auswahl, "0" und Unsinn ergeben alle `null`. */
function numberOrNull(value) {
  if (value === null || value === undefined || value === "") return null;
  const number = Number(value);
  return Number.isFinite(number) && number > 0 ? number : null;
}

/**
 * Zeitpunkt aus einer Antwort des Dienstes.
 *
 * TANSS fuehrt `date` in Unix-SEKUNDEN, sonstige Zeitpunkte kommen als ISO-8601. Beides
 * wird angenommen: eine Sekundenzahl als Millisekunden gelesen ergaebe das Jahr 1970, und
 * das saehe in der Auswahlliste wie ein leeres Feld aus statt wie ein Fehler.
 */
function toDate(value) {
  if (value === null || value === undefined || value === "") return null;
  if (value instanceof Date) return value;
  if (typeof value === "number") return new Date(value < 1e12 ? value * 1000 : value);
  const parsed = new Date(String(value));
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

/** Zeitpunkt in der Schreibweise des Hosts - die Sprache kommt aus dem Betriebssystem. */
function formatMoment(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  try {
    return date.toLocaleString();
  } catch {
    return date.toISOString();
  }
}

/**
 * Dauer als Text, ohne ein Wort dafuer im Programm zu haben.
 *
 * `Intl.NumberFormat` mit der Einheit `minute` liefert die Benennung aus der Laufzeit; ein
 * hartkodiertes "Minuten" waere ein sichtbares Literal ausserhalb von `i18n/de.js`. Faellt
 * die Einheit auf einem aelteren Webview aus, bleibt die nackte Zahl - kein schoener, aber
 * ein ehrlicher Rueckfall.
 */
function formatDuration(minutes) {
  const value = Number(minutes);
  if (!Number.isFinite(value) || value <= 0) return "";
  try {
    return new Intl.NumberFormat(undefined, {
      style: "unit",
      unit: "minute",
      unitDisplay: "long",
    }).format(value);
  } catch {
    return String(value);
  }
}

/** Anzeigetext eines Einsatzortes; unbekannte Werte erscheinen roh statt als Schluessel. */
function locationLabel(code) {
  const key = `location${String(code || "")}`;
  return Object.prototype.hasOwnProperty.call(T.termin, key) ? T.termin[key] : String(code || "");
}

/** Anzeigetext zur gefundenen Kopplungsstufe. */
function sourceLabel(source) {
  switch (source) {
    case "property":
      return T.termin.sourceProperty;
    case "link":
      return T.termin.sourceLink;
    case "syncgroup":
      return T.termin.sourceSyncgroup;
    case "timematch":
      return T.termin.sourceTimematch;
    default:
      return T.termin.sourceNone;
  }
}

/* ------------------------------------------------------------------------ Seite */

/**
 * @param {{root: HTMLElement, me: object, signal: AbortSignal, navigate: Function,
 *          showError: Function, clearError: Function, reload: Function}} ctx
 * @returns {Promise<() => void>} Aufraeumfunktion fuer den Routenwechsel.
 */
export async function render(ctx) {
  ui.replace(ctx.root, ui.loading(T.termin.resolving));

  let appointment;
  let itemKey;
  try {
    appointment = await office.getAppointmentContext();
    // ERST speichern, DANN die Kennung holen - siehe Kopf dieser Datei.
    itemKey = await office.getItemId({ saveFirst: true });
  } catch (error) {
    const code = error && error.code ? error.code : "CLIENT_OFFICE_FAILED";
    ui.replace(
      ctx.root,
      ui.el("div", { class: "stack" }, [
        ui.el("p", { class: "fatal-text", text: errorText({ code }) }),
        ui.button({ label: T.app.retry, onClick: () => void ctx.reload() }),
      ]),
    );
    return () => {};
  }
  if (ctx.signal.aborted) return () => {};

  // Der Instanzschluessel kommt aus `GET /api/me`, das beim Oeffnen des Panes ohnehin
  // geladen wird - NICHT aus `/healthz`. Jene Route ist unauthentifiziert und nennt
  // deshalb bewusst weder Fassung noch Instanz. Stuende der Schluessel hier leer, waere
  // Stufe 1 der Terminkopplung in beide Richtungen tot: Eine am Termin abgelegte
  // Einsatznummer wuerde nie wiedergefunden und nie geschrieben, ohne dass etwas
  // fehlschlaegt - die Maske boete still "Einsatz anlegen" an.
  const defaults = ctx.me && ctx.me.defaults ? ctx.me.defaults : {};

  const state = {
    itemKey,
    instance: ctx.me && ctx.me.instance ? String(ctx.me.instance) : "",
    supportId: null,
    source: "none",
    readOnly: false,
    canCreate: false,
    candidates: [],
    options: null,
    companyId: null,
    companyName: "",
    ticketId: null,
    typeId: numberOrNull(defaults.supportTypeId),
    location: typeof defaults.location === "string" ? defaults.location : "OFFICE",
    internal: defaults.internal === true,
    text: "",
    legacyFound: false,
    busy: false,
  };

  const notices = ui.el("div", { class: "stack" });
  const candidateBox = ui.el("div", { class: "stack" });
  const formBox = ui.el("div", { class: "stack" });
  const companySearchBox = ui.el("div", { class: "stack", props: { hidden: true } });
  const companyResults = ui.el("div", { class: "stack" });
  const statusLine = ui.el("p", { class: "muted" });

  const companySearch = ui.searchField({
    placeholder: T.app.searchPlaceholder,
    minChars: 2,
    onSearch: (term) => void searchCompanies(term),
  });
  ui.mount(companySearchBox, companySearch.root, companyResults);

  /* ------------------------------------------------------- Eigene Eigenschaften */

  /**
   * Liest die am Termin abgelegte Einsatznummer.
   *
   * Nur gueltig, wenn die abgelegte Instanz zu dieser Installation passt: ein aus einer
   * anderen TANSS-Installation stammender Termin darf hier nicht beansprucht werden - die
   * Nummer zeigte dort auf einen voellig anderen Einsatz.
   */
  async function readStoredBinding() {
    try {
      const bag = await office.loadCustomProperties();
      const id = numberOrNull(bag.get(PROP_SUPPORT_ID));
      const instance = String(bag.get(PROP_INSTANCE) || "");
      if (id !== null && instance !== "" && instance === instanceKey()) return id;
    } catch {
      // Kein Zugriff auf die Ablage: die Aufloesung faellt auf die Tabelle des Dienstes
      // und den Zeitfensterabgleich zurueck. Kein Grund abzubrechen.
    }
    return null;
  }

  /**
   * Schluessel dieser TANSS-Installation.
   *
   * Er kommt aus `GET /api/me` und ist derselbe Wert, den der Dienst als
   * `appointment_links.instance` fuehrt - der Hostname der TANSS-API-Basis. Er muss vom
   * Dienst stammen und darf nicht aus der Adresse der Weboberflaeche geraten werden:
   * beide Hosts koennen auseinanderfallen, und eine unter falschem Schluessel abgelegte
   * Einsatznummer wuerde spaeter entweder ignoriert oder - schlimmer - von einer zweiten
   * Installation beansprucht.
   *
   * Ist er nicht zu bekommen, wird die Eigenschaft am Termin NICHT geschrieben: eine
   * Kopplung ohne Instanzangabe ist mehrdeutig, und die massgebliche Kopplung steht
   * ohnehin in TANSS.
   */
  function instanceKey() {
    return state.instance;
  }

  /** Schreibt die Kopplung an den Termin zurueck - Beschleuniger, nie Wahrheit. */
  async function storeBinding(supportId) {
    try {
      if (instanceKey() === "") return;
      const bag = await office.loadCustomProperties();
      bag.set(PROP_SUPPORT_ID, String(supportId));
      bag.set(PROP_INSTANCE, instanceKey());
      // Ohne `save()` ist jedes `set()` wirkungslos - der haeufigste Fehler mit dieser API.
      await bag.save();
    } catch {
      // Die massgebliche Kopplung steht in TANSS; ein fehlgeschlagener Beschleuniger ist
      // kein Grund, einen erfolgreichen Schreibvorgang als Fehler zu melden.
    }
  }

  /* ------------------------------------------------------------------ Aufloesung */

  /**
   * Sucht den Einsatz zum Termin.
   *
   * SONDIERUNG: Stufe 3 der Kopplung (Abgleich ueber `metaInfos.SYNC_GROUP`) braucht die
   * `iCalUId` des Termins. Sie kaeme aus Microsoft Graph
   * (`GET /me/events/{id}?$select=iCalUId`, delegiert `Calendars.Read`); im Taskpane liegt
   * Graph ausschliesslich in `js/mime.js`, und dort gibt es bislang nur den Mailweg. Bis
   * dort ein Kalenderweg steht, geht `icalUid` NICHT mit, und die Aufloesung stuetzt sich
   * auf die Stufen 1 (Eigenschaft am Termin), 2 (Tabelle des Dienstes) und 4
   * (Zeitfensterabgleich). Gegen die echte Instanz zu pruefen: ob `/api/v1/supports/list`
   * `fetchMetaInfos` ueberhaupt annimmt - faellt das aus, ist Stufe 3 ohnehin tot.
   */
  async function resolve() {
    ui.replace(formBox, ui.loading(T.termin.resolving));
    const result = await api.get("api/appointments/resolve", {
      query: {
        // Roh uebergeben: die Normalisierung gehoert in den Dienst, der auch die Tabelle
        // fuehrt - zwei Normalisierungen liefen unweigerlich auseinander.
        itemKey: state.itemKey,
        supportId: state.supportId,
        start: appointment.start instanceof Date ? appointment.start.toISOString() : null,
        durationMinutes: appointment.durationMinutes || null,
      },
      signal: ctx.signal,
    });
    if (ctx.signal.aborted) return;

    if (!result.ok) {
      ctx.showError(result.error);
      ui.replace(
        formBox,
        ui.button({ label: T.app.retry, variant: "ghost", onClick: () => void resolve() }),
      );
      return;
    }

    applyResolution(result.data || {});
    renderNotices();
    renderCandidates();
    await loadOptions();
  }

  /**
   * Uebernimmt die Antwort in den Formularzustand.
   *
   * Reihenfolge der Vorbelegung: gefundener Einsatz schlaegt Altmarker schlaegt Vorgaben
   * aus der Konfiguration. Ein gefundener Einsatz ist eine Tatsache, ein Altmarker eine
   * Absicht von frueher, eine Vorgabe nur eine Vermutung.
   */
  function applyResolution(data) {
    const support = data.support || null;
    state.source = String(data.source || "none");
    state.readOnly = data.readOnly === true;
    state.canCreate = data.canCreate === true;
    state.candidates = Array.isArray(data.candidates) ? data.candidates : [];

    const marker = data.legacyMarker || {};
    const legacy = marker.found === true ? marker.prefill || {} : {};
    // Der Hinweis auf den Altmarker gilt nur, wenn daraus tatsaechlich vorbelegt wurde -
    // steht ein echter Einsatz dahinter, ist der Marker bloss noch Geschichte.
    state.legacyFound = !support && Object.keys(legacy).length > 0;

    if (support) {
      state.supportId = numberOrNull(support.supportId);
      state.companyId = numberOrNull(support.companyId);
      state.companyName = support.companyName || "";
      state.ticketId = numberOrNull(support.ticketId);
      state.typeId = numberOrNull(support.typeId) ?? state.typeId;
      state.location = support.location || state.location;
      state.internal = support.internal === true;
      state.text = typeof support.text === "string" ? support.text : state.text;
      return;
    }

    state.supportId = null;
    state.companyId = numberOrNull(legacy.companyId);
    state.companyName = "";
    state.ticketId = numberOrNull(legacy.ticketId);
    state.typeId = numberOrNull(legacy.typeId) ?? state.typeId;
    if (typeof legacy.location === "string" && legacy.location !== "") {
      state.location = legacy.location;
    }
    if (typeof legacy.isInternal === "boolean") state.internal = legacy.isInternal;
  }

  function renderNotices() {
    const children = [
      ui.banner({ tone: state.supportId ? "ok" : "info", label: sourceLabel(state.source) }),
    ];
    if (state.readOnly) {
      children.push(ui.banner({ tone: "warn", label: T.errors.APPOINTMENT_READONLY }));
    }
    if (state.legacyFound) {
      children.push(ui.banner({ tone: "info", label: T.termin.legacyMarkerFound }));
    }
    if (!state.supportId && !state.canCreate) {
      children.push(ui.banner({ tone: "info", label: T.termin.createDisabled }));
    }
    ui.replace(notices, ...children);
  }

  /** Mehrere Treffer sind kein Fehler, sondern eine Frage an den Benutzer. */
  function renderCandidates() {
    if (state.candidates.length === 0) {
      ui.clear(candidateBox);
      return;
    }
    ui.replace(
      candidateBox,
      ui.banner({ tone: "warn", label: T.termin.candidates }),
      ui.list({
        items: state.candidates,
        renderItem: (candidate) =>
          ui.listRow({
            title: [formatMoment(toDate(candidate.date)), formatDuration(candidate.duration)]
              .filter(Boolean)
              .join(SEPARATOR),
            subtitle: candidate.companyName || "",
            badgeLabel: candidate.ticketId ? String(candidate.ticketId) : "",
            selected: Number(candidate.id) === Number(state.supportId),
            onClick: () => {
              state.supportId = numberOrNull(candidate.id);
              state.candidates = [];
              ui.clear(candidateBox);
              void resolve();
            },
          }),
      }),
    );
  }

  /* --------------------------------------------------------------- Auswahllisten */

  async function loadOptions() {
    const result = await api.get("api/appointments/options", {
      query: { companyId: state.companyId },
      signal: ctx.signal,
    });
    if (ctx.signal.aborted) return;
    if (!result.ok) {
      ctx.showError(result.error);
      ui.replace(
        formBox,
        ui.button({ label: T.app.retry, variant: "ghost", onClick: () => void loadOptions() }),
      );
      return;
    }
    state.options = result.data || {};
    const optionDefaults = state.options.defaults || {};
    if (state.typeId === null) state.typeId = numberOrNull(optionDefaults.supportTypeId);
    if (!state.location && typeof optionDefaults.location === "string") {
      state.location = optionDefaults.location;
    }
    renderForm();
  }

  /* ---------------------------------------------------------------------- Firma */

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
            // Wie in der Ticketmaske: Ein Kunde mit mehreren Standorten fuehrt je
            // Standort eine eigene Kundennummer unter demselben Namen.
            badgeLabel: company.displayId || "",
            tag: company.centralType === "CENTRAL"
              ? { label: T.ticketNeu.companyCentral, tone: "accent" }
              : (company.centralType === "BRANCH"
                ? { label: T.ticketNeu.companyBranch, tone: "muted" }
                : null),
            subtitle: [company.postCode, company.city].filter(Boolean).join(" "),
            onClick: () => {
              state.companyId = numberOrNull(company.id);
              state.companyName = company.name || "";
              // Die Ticketliste der Terminmaske haengt an der Firma; ein Ticket der alten
              // Firma waere nach dem Wechsel eine falsche Zuordnung.
              state.ticketId = null;
              companySearchBox.hidden = true;
              companySearch.input.value = "";
              ui.clear(companyResults);
              void loadOptions();
            },
          }),
      }),
    );
  }

  /* ------------------------------------------------------------------ Formular */

  function renderForm() {
    const options = state.options || {};
    const tickets = Array.isArray(options.tickets) ? options.tickets : [];
    const supportTypes = Array.isArray(options.supportTypes) ? options.supportTypes : [];
    const locations = Array.isArray(options.locations) ? options.locations : [];

    const ticketSelect = ui.select({
      options: tickets.map((ticket) => ({
        value: ticket.id,
        label: `${ticket.id}${SEPARATOR}${ticket.title || ""}`,
      })),
      value: state.ticketId,
      placeholder: T.termin.ticketNone,
      onChange: () => {
        state.ticketId = numberOrNull(ticketSelect.value);
      },
    });

    const typeSelect = ui.select({
      options: supportTypes.map((type) => ({
        value: type.id,
        label: type.name || String(type.id),
      })),
      value: state.typeId,
      placeholder: T.app.none,
      onChange: () => {
        state.typeId = numberOrNull(typeSelect.value);
      },
    });

    const locationSelect = ui.select({
      options: locations.map((code) => ({ value: code, label: locationLabel(code) })),
      value: state.location,
      onChange: () => {
        state.location = locationSelect.value;
      },
    });

    const internalCheck = ui.checkbox({
      label: T.termin.internal,
      checked: state.internal,
      onChange: (event) => {
        state.internal = event.target.checked === true;
      },
    });

    const textInput = ui.textarea({
      value: state.text,
      rows: 5,
      onInput: () => {
        state.text = textInput.value;
      },
    });

    const canSubmit =
      !state.readOnly &&
      state.companyId !== null &&
      state.typeId !== null &&
      (state.supportId !== null || state.canCreate);

    ui.replace(
      formBox,
      ui.row([
        ui.chip({ label: formatMoment(appointment.start), tone: "muted" }),
        ui.chip({ label: formatDuration(appointment.durationMinutes), tone: "muted" }),
      ]),
      ui.el("div", { class: "field" }, [
        ui.el("div", { class: "field-label", text: T.termin.company }),
        ui.row([
          ui.chip({
            label: state.companyName || (state.companyId ? String(state.companyId) : T.app.none),
            tone: state.companyId ? "strong" : "muted",
          }),
          ui.button({
            label: T.ticketNeu.companyChange,
            variant: "ghost",
            disabled: state.readOnly,
            onClick: () => {
              companySearchBox.hidden = !companySearchBox.hidden;
              if (!companySearchBox.hidden) companySearch.focus();
            },
          }),
        ]),
      ]),
      companySearchBox,
      ui.field({ label: T.termin.ticket, control: ticketSelect }),
      ui.field({ label: T.termin.type, control: typeSelect, required: true }),
      ui.field({ label: T.termin.location, control: locationSelect }),
      internalCheck.root,
      ui.field({ label: T.termin.text, control: textInput }),
      statusLine,
      ui.button({
        label: state.supportId ? T.termin.save : T.termin.create,
        variant: "primary",
        disabled: !canSubmit || state.busy,
        onClick: () => void submit(),
      }),
    );
  }

  /* ------------------------------------------------------------- Rueckschreiben */

  /**
   * Schreibt die Felder zurueck.
   *
   * Der Rumpf ist ein GESCHLOSSENES Feldobjekt: was hier nicht steht, geht auch nicht an
   * TANSS. `metaInfos` und `linkedTechnicians` kommen im ganzen Taskpane nicht vor - der
   * Server ERSETZT bei jedem Schreiben den kompletten Satz, ein Teilsatz loescht alles
   * Uebrige, und `linkedTechnicians:[]` loescht die Termine aller weiteren Teilnehmer.
   * Die Firmenzuordnung als Tripel (companyId + linkTypeId + linkId) baut der Dienst.
   */
  async function submit() {
    if (state.busy || state.readOnly) return;
    state.busy = true;
    ctx.clearError();
    ui.setBusy(formBox, true);
    statusLine.textContent = T.app.saving;

    const fields = {
      companyId: state.companyId,
      ticketId: state.ticketId,
      typeId: state.typeId,
      location: state.location,
      internal: state.internal,
      text: state.text,
    };

    const result = state.supportId
      ? await api.put(`api/appointments/${state.supportId}`, {
          json: fields,
          idempotencyKey: api.newIdempotencyKey(),
          signal: ctx.signal,
        })
      : await api.post("api/appointments", {
          json: {
            // Beim Anlegen geht `text` mit: die Notiz ist Teil des Einsatzes, und ohne sie
            // muesste der Benutzer sie unmittelbar nach dem Anlegen ein zweites Mal
            // eintippen. Ein `icalUid` liegt hier nicht vor - siehe `resolve()`.
            ...fields,
            itemKey: state.itemKey,
            start: appointment.start instanceof Date ? appointment.start.toISOString() : null,
            durationMinutes: appointment.durationMinutes || 0,
            title: appointment.subject || "",
          },
          idempotencyKey: api.newIdempotencyKey(),
          signal: ctx.signal,
        });
    if (ctx.signal.aborted) return;

    state.busy = false;
    ui.setBusy(formBox, false);
    statusLine.textContent = "";

    if (!result.ok) {
      if (result.error.code === "APPOINTMENT_READONLY") {
        state.readOnly = true;
        renderNotices();
      }
      renderForm();
      ctx.showError(result.error);
      return;
    }

    const created = state.supportId === null;
    const supportId = numberOrNull(result.data ? result.data.supportId : null) || state.supportId;
    state.supportId = supportId;
    if (supportId !== null) await storeBinding(supportId);
    renderDone(created, supportId, result.data || {});
  }

  function renderDone(created, supportId, data) {
    const ignored = Array.isArray(data.ignoredFields) ? data.ignoredFields : [];
    ui.replace(
      ctx.root,
      ui.el("div", { class: "stack" }, [
        ui.el("h2", {
          class: "section-heading",
          text: created
            ? t("termin.createdHeading", { id: supportId })
            : t("termin.savedHeading", { id: supportId }),
        }),
        ui.banner({ tone: "ok", label: T.app.saved }),
        // TANSS entfernt Felder, die der Anrufer nicht setzen darf, STILL. Deshalb zeigt
        // die Antwort, was wirklich ankam - und die Seite zeigt es weiter.
        ignored.length > 0
          ? ui.el("p", { class: "muted", text: ignored.join(SEPARATOR) })
          : null,
        ui.button({ label: T.app.back, variant: "ghost", onClick: () => void ctx.reload() }),
        ui.button({ label: T.app.closePane, variant: "ghost", onClick: () => office.closePane() }),
      ]),
    );
  }

  /* ------------------------------------------------------------------- Aufbau */

  ui.replace(
    ctx.root,
    notices,
    ui.section({ heading: T.termin.heading, children: [candidateBox, formBox] }),
  );

  state.supportId = await readStoredBinding();
  await resolve();

  // Aufraeumen: der entprellte Zeitgeber der Firmensuche darf nach dem Routenwechsel nicht
  // mehr feuern.
  return () => companySearch.dispose();
}
