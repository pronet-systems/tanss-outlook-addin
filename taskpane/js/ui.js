/**
 * Bausteine der Oberflaeche - die einzige Stelle, an der DOM-Knoten entstehen.
 *
 * Ohne Framework ist handgeschriebener DOM-Code die schwaechste Stelle des Repositories.
 * Dagegen hilft genau eine Regel, und sie wird hier durchgesetzt statt verabredet: Text
 * kommt ausschliesslich ueber `textContent` in den Baum. Die Markup setzenden Eigenschaften
 * des DOM kommen im gesamten Taskpane nicht vor; `tests/test_kein_innerhtml.py` durchsucht
 * `taskpane/` danach und schlaegt fehl, sobald eine auftaucht. Folge sonst: ein Firmenname,
 * ein Ticketbetreff oder eine Absenderadresse aus TANSS wird zu ausfuehrbarem Markup - im
 * selben Origin, der das Microsoft-Token haelt.
 *
 * Die Klassennamen sind der Vertrag zu `app.css`; Seiten setzen keine Stile im Skript.
 */

import { T, errorText, t } from "../i18n/de.js";

/**
 * Eigenschaften, die `el()` direkt am Knoten setzen darf.
 *
 * Eine Erlaubnisliste statt einer Verbotsliste: Was hier nicht steht, wird abgewiesen -
 * auch das, woran beim Schreiben dieser Zeile niemand gedacht hat. Alles Uebrige geht
 * ueber `attrs` (Attribute) oder `text` (Textinhalt), und damit kann kein Markup entstehen.
 */
const ALLOWED_PROPS = new Set([
  "value",
  "checked",
  "disabled",
  "hidden",
  "id",
  "multiple",
  "readOnly",
  "required",
  "rows",
  "selected",
  "tabIndex",
]);

/**
 * Erzeugt ein Element.
 *
 * @param {string} tag
 * @param {{class?: string|string[], text?: string, title?: string, attrs?: object,
 *          props?: object, on?: object, dataset?: object}} [opts]
 * @param {Array<Node|string|null|undefined>} [children]
 */
export function el(tag, opts = {}, children = []) {
  const node = document.createElement(tag);

  if (opts.class) {
    const list = Array.isArray(opts.class) ? opts.class : String(opts.class).split(/\s+/);
    for (const name of list) if (name) node.classList.add(name);
  }
  // Der einzige Weg, wie Text in den Baum kommt.
  if (opts.text !== undefined && opts.text !== null) node.textContent = String(opts.text);
  if (opts.title) node.setAttribute("title", String(opts.title));

  if (opts.attrs) {
    for (const [name, value] of Object.entries(opts.attrs)) {
      if (value === null || value === undefined || value === false) continue;
      node.setAttribute(name, value === true ? "" : String(value));
    }
  }
  if (opts.props) {
    for (const [name, value] of Object.entries(opts.props)) {
      if (!ALLOWED_PROPS.has(name)) {
        throw new Error(`ui.el: Eigenschaft ${name} ist nicht zugelassen`);
      }
      node[name] = value;
    }
  }
  if (opts.dataset) {
    for (const [name, value] of Object.entries(opts.dataset)) node.dataset[name] = String(value);
  }
  if (opts.on) {
    for (const [type, handler] of Object.entries(opts.on)) {
      if (typeof handler === "function") node.addEventListener(type, handler);
    }
  }

  mount(node, ...(Array.isArray(children) ? children : [children]));
  return node;
}

/** Textknoten. Fuer Stellen, an denen kein eigenes Element noetig ist. */
export function text(value) {
  return document.createTextNode(value === null || value === undefined ? "" : String(value));
}

/** Haengt Kinder an; Zeichenketten werden zu Textknoten, `null` wird uebersprungen. */
export function mount(parent, ...children) {
  for (const child of children.flat()) {
    if (child === null || child === undefined || child === false) continue;
    parent.appendChild(child instanceof Node ? child : text(child));
  }
  return parent;
}

/** Leert einen Knoten. `replaceChildren` loest auch die Ereignisbindungen der Kinder. */
export function clear(node) {
  node.replaceChildren();
  return node;
}

/** Ersetzt den Inhalt eines Knotens in einem Zug - ohne Zwischenzustand im Layout. */
export function replace(node, ...children) {
  clear(node);
  return mount(node, ...children);
}

/* ------------------------------------------------------------------ Formular */

export function button({ label, onClick, variant = "", disabled = false, title = "", type = "button" }) {
  return el("button", {
    class: ["btn", variant ? `btn--${variant}` : ""].filter(Boolean),
    text: label,
    title,
    attrs: { type },
    props: { disabled: Boolean(disabled) },
    on: onClick ? { click: onClick } : null,
  });
}

/** Laufende Nummer fuer Feld-Ids; `label for=` braucht eine, und sie muss eindeutig sein. */
let fieldCounter = 0;
function nextId(prefix) {
  fieldCounter += 1;
  return `${prefix}-${fieldCounter}`;
}

/**
 * Beschriftetes Feld. `control` ist ein fertiges Eingabeelement; die Verknuepfung
 * ueber `for`/`id` entsteht hier, damit sie nirgends vergessen wird.
 */
export function field({ label, control, hint = "", required = false }) {
  const id = control.id || nextId("f");
  control.id = id;
  const labelNode = el("label", { class: "field-label", attrs: { for: id } }, [
    text(label),
    required ? el("span", { class: "field-required", text: ` ${T.app.required}` }) : null,
  ]);
  return el("div", { class: "field" }, [
    labelNode,
    control,
    hint ? el("div", { class: "field-hint", text: hint }) : null,
  ]);
}

export function input({ value = "", type = "text", placeholder = "", maxLength = 0, onInput = null, autocomplete = "off" }) {
  return el("input", {
    class: "input",
    attrs: { type, placeholder, autocomplete, maxlength: maxLength > 0 ? maxLength : null },
    props: { value: String(value ?? "") },
    on: onInput ? { input: onInput } : null,
  });
}

export function textarea({ value = "", rows = 6, maxLength = 0, onInput = null, placeholder = "" }) {
  return el("textarea", {
    class: "textarea",
    attrs: { rows, placeholder, maxlength: maxLength > 0 ? maxLength : null },
    props: { value: String(value ?? "") },
    on: onInput ? { input: onInput } : null,
  });
}

/**
 * Auswahlliste. `options` ist `[{value, label, disabled?}]`.
 * Die Beschriftungen kommen ueber `textContent` in die Optionen - auch hier stammen sie
 * aus TANSS und sind damit fremder Text.
 */
export function select({ options = [], value = null, onChange = null, placeholder = "" }) {
  const node = el("select", { class: "select", on: onChange ? { change: onChange } : null });
  if (placeholder) {
    mount(node, el("option", { text: placeholder, attrs: { value: "" } }));
  }
  for (const option of options) {
    mount(
      node,
      el("option", {
        text: option.label,
        attrs: { value: String(option.value), disabled: option.disabled === true },
      }),
    );
  }
  node.value = value === null || value === undefined ? "" : String(value);
  return node;
}

export function checkbox({ label, checked = false, onChange = null }) {
  const box = el("input", {
    class: "checkbox",
    attrs: { type: "checkbox" },
    props: { checked: Boolean(checked) },
    on: onChange ? { change: onChange } : null,
  });
  const id = nextId("c");
  box.id = id;
  return {
    root: el("div", { class: "check" }, [
      box,
      el("label", { class: "check-label", text: label, attrs: { for: id } }),
    ]),
    input: box,
  };
}

/**
 * Zeichenzaehler zu einem Eingabefeld. TANSS schneidet Titel jenseits von 100 Zeichen
 * wortlos ab - der Zaehler macht das sichtbar, bevor es passiert.
 */
export function charCounter(control, max) {
  const node = el("div", { class: "counter" });
  const update = () => {
    const length = String(control.value ?? "").length;
    node.textContent = t(length > max ? "app.charOver" : "app.charCount", { n: length, max });
    node.classList.toggle("counter--over", length > max);
  };
  control.addEventListener("input", update);
  update();
  return node;
}

/* -------------------------------------------------------------------- Anzeige */

/** Kurze Marke, etwa fuer die erkannte Firma samt Herkunftsbegruendung. */
export function chip({ label, title = "", tone = "" }) {
  return el("span", {
    class: ["chip", tone ? `chip--${tone}` : ""].filter(Boolean),
    text: label,
    title,
  });
}

export function badge({ label, tone = "" }) {
  return el("span", { class: ["badge", tone ? `badge--${tone}` : ""].filter(Boolean), text: label });
}

/** Schluessel-Wert-Liste; die Diagnoseseite besteht fast nur daraus. */
export function kv(pairs) {
  const node = el("dl", { class: "kv" });
  for (const [key, value] of pairs) {
    if (value === undefined) continue;
    mount(
      node,
      el("dt", { class: "kv-key", text: key }),
      el("dd", { class: "kv-val", text: value === null || value === "" ? T.app.none : value }),
    );
  }
  return node;
}

export function spinner() {
  return el("span", { class: "spinner", attrs: { "aria-hidden": "true" } });
}

/** Ladezustand mit Text. Ein Pane ohne Text sieht aus wie ein haengendes Pane. */
export function loading(label = T.app.loading) {
  return el("div", { class: "loading", attrs: { role: "status" } }, [
    spinner(),
    el("span", { class: "loading-text", text: label }),
  ]);
}

/**
 * Hinweisband mit optionaler Handlung ("Stattdessen anhaengen", "Mail jetzt anhaengen").
 * `tone` ist eine von: info, warn, ok.
 */
export function banner({ tone = "info", label, actionLabel = "", onAction = null }) {
  return el("div", { class: ["banner", `banner--${tone}`] }, [
    el("span", { class: "banner-text", text: label }),
    actionLabel && onAction
      ? button({ label: actionLabel, onClick: onAction, variant: "ghost" })
      : null,
  ]);
}

/**
 * Fehlerleiste am Kopf einer Seite.
 *
 * Zeigt den fertigen Satz des Dienstes und - klein - die Vorgangsnummer, unter der der
 * Administrator denselben Fall im Journal wiederfindet. Ohne diese Nummer bleibt bei einer
 * Rueckfrage nur die Uhrzeit.
 */
export function errorBar() {
  const message = el("span", { class: "errbar-msg" });
  const trace = el("span", { class: "errbar-trace" });
  const root = el(
    "div",
    // `hidden` als Eigenschaft, nicht als Attribut: so steht der Anfangszustand auch dann
    // richtig, wenn niemand auf die Rueckspiegelung Attribut -> Eigenschaft baut.
    { class: "errbar", attrs: { role: "alert" }, props: { hidden: true } },
    [
      el("div", { class: "errbar-body" }, [message, trace]),
      el("button", {
        class: "errbar-close",
        text: "×",
        title: T.app.close,
        attrs: { type: "button", "aria-label": T.app.close },
        on: { click: () => hide() },
      }),
    ],
  );

  function show(error) {
    // Die Auswahl des Textes liegt in i18n, damit hier kein sichtbares Literal steht.
    message.textContent = errorText(error);
    const id = error && (error.traceId || (error.detail && error.detail.traceId));
    trace.textContent = id ? t("app.operationId", { id }) : "";
    root.hidden = false;
  }

  function hide() {
    root.hidden = true;
    message.textContent = "";
    trace.textContent = "";
  }

  return { root, show, hide, get visible() { return root.hidden === false; } };
}

/**
 * Liste mit Leermeldung. `renderItem` liefert den Knoten je Eintrag; die Leermeldung
 * ist Pflicht, weil eine leere Liste ohne Text wie ein Ladefehler aussieht.
 */
export function list({ items, renderItem, emptyLabel = T.app.noResults }) {
  if (!items || items.length === 0) {
    return el("div", { class: "list-empty", text: emptyLabel });
  }
  const node = el("ul", { class: "list" });
  for (const item of items) {
    mount(node, el("li", { class: "list-item" }, [renderItem(item)]));
  }
  return node;
}

/** Eine anklickbare Zeile: Titel, Nebenzeile, optionale Marke rechts. */
/**
 * Eine Zeile einer Trefferliste.
 *
 * `tag` ist ein ZWEITES Abzeichen und tritt bewusst vor das erste. Zwei gleich aussehende
 * Abzeichen nebeneinander waeren zwei Dinge von gleichem Gewicht - deshalb traegt `tag`
 * eine Farbe und `badgeLabel` bleibt schlicht: Das eine benennt eine Art, das andere ist
 * eine Kennung. Wer beides gleich auszeichnet, zwingt zum Lesen, wo ein Blick genuegen
 * sollte.
 */
export function listRow({ title, subtitle = "", badgeLabel = "", tag = null,
  onClick = null, selected = false }) {
  return el(
    "button",
    {
      class: ["row-btn", selected ? "row-btn--selected" : ""].filter(Boolean),
      attrs: { type: "button" },
      on: onClick ? { click: onClick } : null,
    },
    [
      el("span", { class: "row-main" }, [
        el("span", { class: "row-title", text: title }),
        subtitle ? el("span", { class: "row-sub", text: subtitle }) : null,
      ]),
      tag && tag.label ? badge({ label: tag.label, tone: tag.tone || "" }) : null,
      badgeLabel ? badge({ label: badgeLabel }) : null,
    ],
  );
}

/* ------------------------------------------------------------------ Verhalten */

/**
 * Entprellung mit abbrechbarem Nachlauf.
 *
 * Ohne Entprellung feuert jeder Tastenanschlag eine Suche ab. Ohne `cancel()` laeuft
 * ausserdem nach dem Verlassen der Seite noch eine Anfrage in eine abgeraeumte
 * Oberflaeche - und schreibt in Knoten, die es nicht mehr gibt.
 */
export function debounce(fn, delay = 300) {
  let timer = 0;
  const wrapped = (...args) => {
    window.clearTimeout(timer);
    timer = window.setTimeout(() => fn(...args), delay);
  };
  wrapped.cancel = () => window.clearTimeout(timer);
  wrapped.flush = (...args) => {
    window.clearTimeout(timer);
    fn(...args);
  };
  return wrapped;
}

/**
 * Entprelltes Suchfeld.
 *
 * `onSearch(begriff)` wird erst nach `delay` Millisekunden und erst ab `minChars` Zeichen
 * gerufen; ein geleertes Feld meldet einmal den leeren Begriff, damit die Trefferliste
 * verschwindet statt stehenzubleiben.
 */
export function searchField({
  placeholder = T.app.searchPlaceholder,
  minChars = 2,
  delay = 300,
  onSearch,
  value = "",
}) {
  const hint = el("div", { class: "search-hint" });
  const control = el("input", {
    class: "input search-input",
    attrs: { type: "search", placeholder, autocomplete: "off", "aria-label": T.app.search },
    props: { value: String(value ?? "") },
  });

  const run = debounce((term) => onSearch(term), delay);

  control.addEventListener("input", () => {
    const term = control.value.trim();
    if (term === "") {
      hint.textContent = "";
      run.flush("");
      return;
    }
    if (term.length < minChars) {
      hint.textContent = t("app.searchTooShort", { min: minChars });
      run.cancel();
      return;
    }
    hint.textContent = "";
    run(term);
  });

  // Enter sucht sofort - wer die Ticketnummer kennt, will nicht auf die Entprellung warten.
  control.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      const term = control.value.trim();
      if (term.length >= minChars || term === "") run.flush(term);
    }
  });

  const root = el("div", { class: "search" }, [control, hint]);

  return {
    root,
    input: control,
    focus: () => control.focus(),
    setBusy: (busy) => root.classList.toggle("search--busy", Boolean(busy)),
    dispose: () => run.cancel(),
    get value() {
      return control.value.trim();
    },
  };
}

/** Sperrt einen Bereich waehrend eines laufenden Vorgangs - sichtbar und fuer die Tastatur. */
export function setBusy(node, busy) {
  node.classList.toggle("is-busy", Boolean(busy));
  node.setAttribute("aria-busy", busy ? "true" : "false");
  for (const control of node.querySelectorAll("button, input, select, textarea")) {
    control.disabled = Boolean(busy);
  }
}

/** Abschnitt mit Ueberschrift. Haelt die Seiten flach und die Ueberschriftenebene stabil. */
/**
 * Eine Ziehharmonika: zugeklappter Bereich mit Kopfzeile.
 *
 * Gebaut aus `details`/`summary` und nicht aus einem eigenen Knopf mit Umschaltlogik. Der
 * Browser bringt das Auf- und Zuklappen, die Tastaturbedienung und die Ansage an
 * Vorleseprogramme mit; jede Eigenbau-Loesung muesste das nachbilden und liesse
 * erfahrungsgemaess die Haelfte davon weg.
 *
 * `summary` nimmt einen Knoten, keinen blossen Text: Die Kopfzeile soll mitsagen, was im
 * zugeklappten Bereich steht. Eine Ziehharmonika, die ihren Inhalt verschweigt, ist
 * schlechter als gar keine - der Benutzer muss sie oeffnen, nur um zu sehen, ob er sie
 * oeffnen wollte.
 */
export function collapsible({ summary, children = [], open = false }) {
  const kopf = el("summary", { class: "fold-head" });
  mount(kopf, summary);
  // `open` geht ueber `attrs` und nicht ueber `props`: Es ist ein boolesches ATTRIBUT,
  // und genau dafuer ist `attrs` da - `false` laesst es weg, `true` setzt es leer. Der
  // Weg ueber `props` waere abgewiesen worden, und zu Recht: Die Erlaubnisliste dort
  // haelt fern, woran beim Schreiben dieser Zeile niemand gedacht hat.
  const box = el("details", { class: "fold", attrs: { open: open === true } });
  mount(box, kopf, ...children);
  return box;
}

export function section({ heading = "", children = [] }) {
  return el("section", { class: "section" }, [
    heading ? el("h2", { class: "section-heading", text: heading }) : null,
    ...children,
  ]);
}

/** Senkrechter Stapel mit einheitlichem Abstand - ersetzt Abstands-Stile im Skript. */
export function stack(children = []) {
  return el("div", { class: "stack" }, children);
}

/** Waagerechte Reihe, die auf schmalen Panes umbricht. */
export function row(children = []) {
  return el("div", { class: "row" }, children);
}
