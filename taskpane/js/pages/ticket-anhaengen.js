/**
 * Seite `#/ticket/anhaengen`: die geoeffnete E-Mail an ein BESTEHENDES Ticket haengen.
 *
 * Der kleinste vollstaendige Arbeitsablauf des Add-ins und deshalb der wichtigste: acht
 * Mails an demselben Vorgang sind der Regelfall, nicht acht Tickets.
 *
 * Drei Ausgaenge sind KEIN Fehler, sondern eigene Zustaende, und jeder bekommt hier seine
 * eigene Anzeige:
 *
 * - `MAIL_ALREADY_ATTACHED` - die Dublettenpruefung des Dienstes hat angeschlagen. TANSS
 *   selbst prueft auf keinem Anhaengepfad auf die Message-ID; ohne diese Warnung haengt
 *   dieselbe Mail nach dem zweiten Klick zweimal am Ticket, mit zweiter Benachrichtigung
 *   an den zugewiesenen Techniker. Es ist eine WARNUNG, keine Sperre: wer es trotzdem will,
 *   bekommt "Trotzdem anhaengen".
 * - `MAIL_STORED_AS_DOCUMENT` - TANSS konnte die Nachricht nicht als Mail auswerten und hat
 *   sie still als gewoehnliches Dokument abgelegt. Sie ist da, aber in der Mailansicht des
 *   Tickets nicht zu finden. Als Erfolg gemeldet waere das die unangenehmste Ueberraschung,
 *   die dieses Werkzeug produzieren kann.
 * - `MAIL_TOO_LARGE` faellt in der Regel schon im Taskpane an, bevor etwas uebertragen wird.
 *
 * Die Suche ist entprellt und laeuft ueber `ui.searchField` - ohne das feuerte jeder
 * Tastenanschlag eine Anfrage ab.
 */

import { T, errorText, formatBytes, t } from "../../i18n/de.js";
import * as api from "../api.js";
import * as mime from "../mime.js";
import * as office from "../office.js";
import * as ui from "../ui.js";

/** Gemerkter Suchbereich (`POST /api/prefs`, geschlossene Schluesselliste). */
const PREF_SCOPE = "searchScope";

const SCOPE_COMPANY = "company";
const SCOPE_ALL = "all";

/** Trennzeichen zwischen zwei Angaben einer Zeile. Satzzeichen, kein uebersetzbarer Text. */
const SEPARATOR = " · ";

/**
 * Adresse des Tickets in der TANSS-Weboberflaeche.
 *
 * Ueber `URL`/`URLSearchParams` statt per Zeichenkettenverkettung, und die Ticketnummer
 * geht nur als GEPRUEFTE Zahl hinein. Werden Nummer und Titel unkodiert verkettet,
 * erzeugt ein Titel mit `&` oder `#` eine Adresse, die woandershin zeigt.
 */
function ticketUrl(frontendBase, ticketId) {
  const id = Number(ticketId);
  if (!frontendBase || !Number.isInteger(id) || id <= 0) return "";
  try {
    const base = String(frontendBase).endsWith("/") ? String(frontendBase) : `${frontendBase}/`;
    const url = new URL("index.php", base);
    url.searchParams.set("section", "bug");
    url.searchParams.set("sub", "view");
    url.searchParams.set("bugID", String(id));
    return url.toString();
  } catch {
    return "";
  }
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

  // Derselbe Aufruf wie in der Ticketmaske: er liefert die Firma des Absenders - und damit
  // den Suchbereich - sowie die Tickets, an denen diese Mail schon haengt.
  const contextResult = await api.post("api/mail/context", {
    json: {
      senderAddress: message.from ? message.from.address : "",
      senderName: message.from ? message.from.name : "",
      subject: message.subject,
      internetMessageId: message.internetMessageId,
    },
    signal: ctx.signal,
  });
  if (ctx.signal.aborted) return () => {};

  const data = contextResult.ok ? contextResult.data || {} : {};
  if (!contextResult.ok) {
    // Ohne Absenderaufloesung bleibt die Suche ueber alle Firmen - das ist eine
    // Einschraenkung, kein Abbruch. Der Fehler steht in der Leiste, die Maske arbeitet.
    ctx.showError(contextResult.error);
  }

  const prefs = ctx.me && ctx.me.prefs ? ctx.me.prefs : {};
  const company = data.company && data.company.id ? data.company : null;

  const state = {
    company,
    scope: company && prefs[PREF_SCOPE] !== SCOPE_ALL ? SCOPE_COMPANY : SCOPE_ALL,
    includeDone: false,
    /** Ticketnummern, an denen diese Mail laut Dublettentabelle schon haengt. */
    attachedTo: new Set(
      (Array.isArray(data.alreadyAttachedTo) ? data.alreadyAttachedTo : []).map((entry) =>
        Number(entry.ticketId),
      ),
    ),
    selected: null,
    term: "",
    busy: false,
    /** Letztes Suchergebnis - damit das Anwaehlen einer Zeile keine Anfrage kostet. */
    items: [],
    recent: [],
    tooMany: false,
  };

  const notices = ui.el("div", { class: "stack" });
  const selectedBox = ui.el("div", { class: "stack" });
  const resultsBox = ui.el("div", { class: "stack" });
  const statusLine = ui.el("p", { class: "muted" });

  const search = ui.searchField({
    placeholder: T.ticketAnhaengen.searchPlaceholder,
    minChars: 2,
    onSearch: (term) => {
      state.term = term;
      void loadTickets();
    },
  });

  const scopeSelect = ui.select({
    options: [
      { value: SCOPE_COMPANY, label: T.ticketAnhaengen.scopeCompany, disabled: company === null },
      { value: SCOPE_ALL, label: T.ticketAnhaengen.scopeAll },
    ],
    value: state.scope,
    onChange: () => {
      state.scope = scopeSelect.value === SCOPE_COMPANY ? SCOPE_COMPANY : SCOPE_ALL;
      // Bewusst ohne Abbruchsignal: das Merken ueberdauert die Seite, nicht umgekehrt.
      void api.post("api/prefs", { json: { key: PREF_SCOPE, value: state.scope } });
      void loadTickets();
    },
  });

  const doneCheck = ui.checkbox({
    label: T.ticketAnhaengen.includeDone,
    checked: state.includeDone,
    onChange: (event) => {
      state.includeDone = event.target.checked === true;
      void loadTickets();
    },
  });

  /* ---------------------------------------------------------------- Trefferliste */

  /**
   * Holt Treffer oder - ohne Suchbegriff - die offenen Tickets der Firma samt der zuletzt
   * benutzten. Beides kommt aus derselben Route, damit es nur eine Fehlerbehandlung gibt.
   */
  async function loadTickets() {
    search.setBusy(true);
    ui.replace(resultsBox, ui.loading(state.term === "" ? T.app.loading : T.app.searching));

    const result = await api.get("api/search/tickets", {
      query: {
        q: state.term,
        companyId: state.scope === SCOPE_COMPANY && state.company ? state.company.id : null,
        scope: state.scope,
        includeDone: state.includeDone ? "true" : "false",
        limit: 25,
        // Ohne die Message-Id kann der Dienst `alreadyHasThisMail` nicht ableiten: seine
        // Dublettentabelle ist auf deren Hash geschluesselt. Fehlte sie, waere die
        // Kennzeichnung bei jedem Treffer falsch-negativ.
        internetMessageId: message.internetMessageId,
      },
      signal: ctx.signal,
    });
    if (ctx.signal.aborted) return;
    search.setBusy(false);

    if (!result.ok) {
      ctx.showError(result.error);
      ui.replace(
        resultsBox,
        ui.button({ label: T.app.retry, variant: "ghost", onClick: () => void loadTickets() }),
      );
      return;
    }

    state.items = Array.isArray(result.data.items) ? result.data.items : [];
    state.recent = Array.isArray(result.data.recent) ? result.data.recent : [];
    state.tooMany = result.data.tooMany === true;
    renderResults();
  }

  /** Zeichnet die Liste aus dem letzten Ergebnis - ohne den Dienst noch einmal zu fragen. */
  function renderResults() {
    ui.replace(
      resultsBox,
      state.tooMany ? ui.banner({ tone: "warn", label: T.app.tooMany }) : null,
      state.term === "" && state.recent.length > 0
        ? ui.section({
            heading: T.ticketAnhaengen.recent,
            children: [ui.list({ items: state.recent, renderItem: ticketRow })],
          })
        : null,
      ui.section({
        heading: T.ticketAnhaengen.results,
        children: [ui.list({ items: state.items, renderItem: ticketRow })],
      }),
    );
  }

  /** Eine Trefferzeile: Nummer und Titel oben, Firma unten, Status als Marke rechts. */
  function ticketRow(ticket) {
    const known = ticket.alreadyHasThisMail === true || state.attachedTo.has(Number(ticket.id));
    const subtitle = [ticket.companyName || "", known ? T.ticketAnhaengen.alreadyAttached : ""]
      .filter(Boolean)
      .join(SEPARATOR);
    return ui.listRow({
      title: `${ticket.id}${SEPARATOR}${ticket.title || ""}`,
      subtitle,
      badgeLabel: ticket.statusName || "",
      selected: state.selected !== null && Number(state.selected.id) === Number(ticket.id),
      onClick: () => {
        state.selected = ticket;
        ctx.clearError();
        renderSelected();
        renderResults();
      },
    });
  }

  /* ------------------------------------------------------------------- Auswahl */

  function renderSelected() {
    if (!state.selected) {
      ui.clear(selectedBox);
      return;
    }
    const ticket = state.selected;
    const known = ticket.alreadyHasThisMail === true || state.attachedTo.has(Number(ticket.id));

    ui.replace(
      selectedBox,
      ui.section({
        children: [
          ui.row([
            ui.chip({
              label: `${ticket.id}${SEPARATOR}${ticket.title || ""}`,
              title: ticket.companyName || "",
              tone: "strong",
            }),
          ]),
          known
            ? ui.banner({
                tone: "warn",
                label: T.ticketAnhaengen.alreadyAttached,
                actionLabel: T.ticketAnhaengen.forceLabel,
                onAction: () => void attach(true),
              })
            : null,
          ui.el("p", { class: "muted", text: attachmentLabel() }),
          mime.currentSource() === "officejs"
            ? ui.el("p", { class: "muted", text: T.mail.reconstructed })
            : null,
          statusLine,
          ui.button({
            label: T.ticketAnhaengen.attach,
            variant: "primary",
            disabled: state.busy || known,
            onClick: () => void attach(false),
          }),
        ],
      }),
    );
  }

  /** Anzahl echter Anhaenge; eingebettete Bilder sind fuer den Benutzer keine. */
  function attachmentLabel() {
    let count = 0;
    try {
      count = office.getAttachments().filter((attachment) => !attachment.isInline).length;
    } catch {
      count = 0;
    }
    if (count === 0) return T.mail.attachmentsNone;
    if (count === 1) return T.mail.attachmentsOne;
    return t("mail.attachmentsMany", { n: count });
  }

  /* ------------------------------------------------------------------ Ablegen */

  /**
   * Haengt die Mail an das gewaehlte Ticket.
   *
   * `force` ueberspringt NUR die lokale Dublettenwarnung; alles andere - Groessengrenze,
   * Auswertbarkeit, Rechte - prueft der Dienst unveraendert.
   */
  async function attach(force) {
    if (!state.selected || state.busy) return;
    const ticketId = Number(state.selected.id);
    state.busy = true;
    ctx.clearError();
    ui.setBusy(selectedBox, true);
    statusLine.textContent = T.mail.fetching;

    const limit = ctx.me && ctx.me.limits ? Number(ctx.me.limits.maxEmlBytes) || 0 : 0;
    const mail = await mime.fetchMessageMime({ maxBytes: limit });
    if (ctx.signal.aborted) return;
    if (!mail.ok) {
      finish();
      ctx.showError(mail.error);
      return;
    }
    for (const warning of mail.data.warnings) {
      ui.mount(notices, ui.banner({ tone: "warn", label: warning }));
    }
    statusLine.textContent = `${t("mail.sizeInfo", { size: formatBytes(mail.data.bytes) })}`;

    const form = new FormData();
    form.set("eml", mail.data.blob, mail.data.filename);
    form.set("internetMessageId", mail.data.internetMessageId || "");
    form.set("subject", mail.data.subject || "");
    if (force) form.set("force", "true");

    statusLine.textContent = T.ticketAnhaengen.attaching;
    const result = await api.post(`api/tickets/${ticketId}/mail`, {
      form,
      // Neuer Schluessel je Versuch: `force` aendert den Inhalt, und derselbe Schluessel
      // mit anderem Inhalt ergibt CONFLICT statt der gewuenschten Wiederholung.
      idempotencyKey: api.newIdempotencyKey(),
      signal: ctx.signal,
    });
    if (ctx.signal.aborted) return;
    finish();

    if (result.ok) {
      renderDone(ticketId, result.data || {});
      return;
    }
    handleFailure(ticketId, result.error);
  }

  function finish() {
    state.busy = false;
    ui.setBusy(selectedBox, false);
    statusLine.textContent = "";
    renderSelected();
  }

  /**
   * Die beiden Faelle, die keine gewoehnlichen Fehler sind.
   *
   * Bei `MAIL_STORED_AS_DOCUMENT` liegt die Mail in TANSS - nur eben als Dokument statt als
   * Mail. Der Benutzer bekommt deshalb den Weg zum Ticket und keinen Wiederholungsknopf:
   * ein zweiter Versuch endete genauso.
   */
  function handleFailure(ticketId, error) {
    const code = error && error.code;
    if (code === "MAIL_ALREADY_ATTACHED") {
      // Der Satz des Dienstes nennt Ticket und Zeitpunkt; die Handlung dazu steht genau
      // einmal, naemlich am gewaehlten Ticket weiter unten.
      state.attachedTo.add(ticketId);
      ui.replace(
        notices,
        ui.banner({ tone: "warn", label: error.message || T.ticketAnhaengen.alreadyAttached }),
      );
      renderSelected();
      return;
    }
    if (code === "MAIL_STORED_AS_DOCUMENT") {
      const url = ticketUrl(ctx.me ? ctx.me.tanssFrontendUrl : "", ticketId);
      ui.replace(
        ctx.root,
        ui.el("div", { class: "stack" }, [
          ui.banner({ tone: "warn", label: error.message || T.errors.MAIL_STORED_AS_DOCUMENT }),
          url
            ? ui.button({
                label: T.app.openInTanss,
                variant: "primary",
                onClick: () => {
                  if (!office.openInBrowser(url)) ctx.showError({ code: "CLIENT_OFFICE_FAILED" });
                },
              })
            : null,
          ui.button({ label: T.app.back, variant: "ghost", onClick: () => void ctx.reload() }),
        ]),
      );
      return;
    }
    ctx.showError(error);
  }

  function renderDone(ticketId, result) {
    const url = ticketUrl(ctx.me ? ctx.me.tanssFrontendUrl : "", ticketId);
    ui.replace(
      ctx.root,
      ui.el("div", { class: "stack" }, [
        ui.el("h2", {
          class: "section-heading",
          text: t("ticketAnhaengen.successHeading", { id: ticketId }),
        }),
        ui.banner({
          tone: "ok",
          label: result.storedAsDocument === true
            ? T.errors.MAIL_STORED_AS_DOCUMENT
            : T.ticketNeu.successMailAttached,
        }),
        result.bytes
          ? ui.el("p", {
              class: "muted",
              text: t("mail.sizeInfo", { size: formatBytes(result.bytes) }),
            })
          : null,
        url
          ? ui.button({
              label: T.app.openInTanss,
              variant: "primary",
              onClick: () => {
                if (!office.openInBrowser(url)) ctx.showError({ code: "CLIENT_OFFICE_FAILED" });
              },
            })
          : null,
        ui.button({ label: T.app.closePane, variant: "ghost", onClick: () => office.closePane() }),
      ]),
    );
  }

  /* ------------------------------------------------------------------- Aufbau */

  ui.replace(
    ctx.root,
    notices,
    ui.section({
      heading: T.ticketAnhaengen.heading,
      children: [
        search.root,
        company ? ui.field({ label: T.ticketNeu.company, control: scopeSelect }) : null,
        doneCheck.root,
      ],
    }),
    selectedBox,
    resultsBox,
  );

  await loadTickets();

  // Aufraeumen: der entprellte Zeitgeber des Suchfeldes darf nach dem Routenwechsel nicht
  // mehr feuern - er schriebe sonst in Knoten, die es nicht mehr gibt.
  return () => search.dispose();
}
