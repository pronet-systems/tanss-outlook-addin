/**
 * Alle sichtbaren Texte des Taskpanes.
 *
 * Es gibt genau eine Sprache (Deutsch) und trotzdem diese Datei: nicht wegen einer
 * spaeteren Uebersetzung, sondern weil verstreute Literale unweigerlich auseinanderlaufen.
 * Stehen dieselbe Meldung an vier Stellen und drei Fassungen davon widersprechen einander,
 * kann niemand mehr sagen, welche der Benutzer zu sehen bekommt.
 *
 * Regel: Im uebrigen Taskpane steht kein Literal, das ein Benutzer lesen kann.
 * Wer hier nichts findet, legt hier etwas an - nicht dort, wo es gebraucht wird.
 *
 * Fehlertexte: Der Dienst liefert in `error.message` immer einen fertigen deutschen Satz.
 * Die Tabelle `errors` ist der Rueckfall fuer den Fall, in dem gar kein Koerper ankommt
 * (Netzabbruch, Proxy-Fehlerseite, 502 ohne JSON) - dann darf das Pane nicht stumm bleiben.
 */

/** Tiefes Einfrieren. Ein versehentliches `T.app.save = ...` faellt sonst erst beim Benutzer auf. */
function deepFreeze(value) {
  if (value && typeof value === "object" && !Object.isFrozen(value)) {
    Object.freeze(value);
    for (const key of Object.keys(value)) deepFreeze(value[key]);
  }
  return value;
}

export const T = deepFreeze({
  app: {
    name: "TANSS",
    /**
     * Alternativtext des Herausgeberzeichens im Kopf.
     *
     * Er steht hier und nicht als leeres `alt`: Ein Bild ohne Alternativtext ist fuer
     * einen Screenreader unsichtbar, und gerade auf der Anmeldeseite soll hoerbar sein,
     * wessen Oberflaeche nach dem Kennwort fragt.
     */
    vendor: "ProNet Systems",
    loading: "Wird geladen …",
    working: "Einen Moment …",
    retry: "Erneut versuchen",
    cancel: "Abbrechen",
    close: "Schließen",
    back: "Zurück",
    save: "Speichern",
    saving: "Wird gespeichert …",
    saved: "Gespeichert.",
    create: "Anlegen",
    search: "Suchen",
    searchPlaceholder: "Suchbegriff …",
    searchTooShort: "Mindestens {min} Zeichen eingeben.",
    searching: "Wird gesucht …",
    noResults: "Keine Treffer.",
    tooMany: "Zu viele Treffer — bitte den Suchbegriff verfeinern.",
    optional: "optional",
    required: "Pflichtfeld",
    none: "—",
    unknown: "unbekannt",
    yes: "ja",
    no: "nein",
    charCount: "{n} von {max} Zeichen",
    charOver: "{n} von {max} Zeichen — wird auf {max} gekürzt",
    bytes: "{n} Bytes",
    kilobytes: "{n} kB",
    megabytes: "{n} MB",
    openInTanss: "In TANSS öffnen",
    closePane: "Fenster schließen",
    detailsShow: "Einzelheiten anzeigen",
    detailsHide: "Einzelheiten ausblenden",
    operationId: "Vorgang {id}",
  },

  nav: {
    ticketNeu: "Ticket erstellen",
    ticketAnhaengen: "An Ticket anhängen",
    termin: "TANSS-Einsatz",
    anmeldung: "Anmeldung",
    diagnose: "Diagnose",
    unknownRoute: "Diese Ansicht gibt es nicht.",
    unknownRouteHint:
      "Der Aufruf zeigt auf eine Seite, die es in dieser Fassung nicht gibt. " +
      "Wahrscheinlich ist das Manifest neuer als das Taskpane oder umgekehrt.",
  },

  boot: {
    officeWaiting: "Outlook wird abgewartet …",
    hostCheck: "Fähigkeiten des Hosts werden geprüft …",
    signingIn: "Anmeldung läuft …",
    profileLoading: "Zuordnung zum TANSS-Mitarbeiter wird geholt …",
    ready: "Bereit.",
  },

  /** Der Host kann weniger, als dieses Add-in braucht. Jede Meldung nennt die Folge. */
  host: {
    notOutlook:
      "Dieses Add-in läuft nur in Outlook. Der aktuelle Host meldet keine Postfachschnittstelle.",
    officeSilent:
      "Outlook hat sich nicht gemeldet. Die Office-Bibliothek ist geladen, aber die Verbindung " +
      "zum Host kam nicht zustande — ohne sie kann das Add-in weder die E-Mail noch den Termin " +
      "lesen. Ein erneuter Versuch hilft oft; sonst zeigt die Diagnose, was der Host meldet.",
    setTooOld:
      "Dieser Outlook-Stand ist zu alt: er meldet das Anforderungspaket Mailbox {required} nicht. " +
      "Ohne dieses Paket kann das Add-in nicht erkennen, ob das Postfach in Exchange Online " +
      "oder auf einem lokalen Exchange liegt — und damit auch nicht, wie es die Mail holen soll.",
    attachmentsUnsupported:
      "Dieser Outlook-Stand meldet Mailbox 1.8 nicht. Anhänge können deshalb nicht mitgegeben " +
      "werden; die Mail würde ohne ihre Anhänge in TANSS landen.",
    wrongItemMessage: "Für diese Ansicht muss eine E-Mail geöffnet sein.",
    wrongItemAppointment: "Für diese Ansicht muss ein Termin geöffnet sein.",
    noItem: "Es ist kein Element geöffnet, auf das sich das Add-in beziehen könnte.",
  },

  auth: {
    targetHost: "Anmeldung bei {host}",
    instanceHeading: "TANSS-Adresse",
    instanceMissing: "Für dieses Add-in ist noch keine TANSS-Adresse hinterlegt.",
    instanceLabel: "Adresse der TANSS-API",
    instanceHint:
      "Meist die Adresse Ihrer TANSS-Oberfläche, ergänzt um /backend. "
      + "Sie wird auf diesem Rechner gemerkt.",
    instanceSave: "Übernehmen",
    instanceRejected:
      "Diese Adresse wird nicht angenommen. Zulässig ist nur eine https-Adresse — und "
      + "nur eine, die für dieses Add-in freigegeben ist.",

    signedInAs: "Angemeldet als {name}",
    employeeId: "TANSS-Mitarbeiter {id}",
    signIn: "Anmelden",
    signInTanss: "Bei TANSS anmelden",
    username: "Benutzername",
    password: "Kennwort",
    twoFactor: "Zweiter Faktor",
    twoFactorHint: "Nur ausfüllen, wenn TANSS danach fragt.",
    fallbackNotice:
      "Die Anmeldung erfolgt mit Ihren TANSS-Zugangsdaten. Sie gilt auf diesem Rechner " +
      "und wird selbsttätig erneuert; eine neue Anmeldung steht erst nach längerer Zeit " +
      "wieder an.",
    methodNaa: "Microsoft-Konto (NAA)",
    methodTanss: "TANSS-Anmeldung",
    methodNone: "nicht angemeldet",
    signOut: "Abmelden",
  },

  mail: {
    fetching: "E-Mail wird aus dem Postfach geholt …",
    fetchingGraph: "E-Mail wird über Microsoft Graph geholt …",
    fetchingOfficejs: "E-Mail wird aus Outlook zusammengesetzt …",
    attachmentsOne: "1 Anhang wird mitgegeben.",
    attachmentsMany: "{n} Anhänge werden mitgegeben.",
    attachmentsNone: "Diese Mail hat keine Anhänge.",
    reconstructed:
      "Diese Mail wird aus Outlook rekonstruiert, nicht im Original übernommen: " +
      "Empfangskette, DKIM-Signatur und Bezugsverweise fehlen.",
    cloudAttachmentSkipped:
      "Der Anhang „{name}“ liegt nur als Cloud-Verweis vor und kann nicht mitgegeben werden.",
    attachmentFailed: "Der Anhang „{name}“ ließ sich nicht lesen und fehlt deshalb.",
    sizeInfo: "Größe: {size}",
    tooLargeLocal:
      "Die Mail ist mit {size} größer als die zugelassenen {limit}. Sie wird nicht übertragen.",
  },

  ticketNeu: {
    heading: "Ticket aus dieser E-Mail",
    company: "Firma",
    companyFromAddress: "aus der Absenderadresse erkannt",
    companyFromDomain: "aus der Absenderdomäne erkannt",
    companyUnknown: "nicht erkannt — bitte auswählen",
    companyChange: "Andere Firma wählen",
    remitter: "Melder",
    companyFirst:
      "Bitte zuerst die Firma wählen. Ohne sie ginge die Suche über alle Kunden, und der " +
      "gewählte Ansprechpartner gehörte am Ende zu keinem davon.",
    remitterUnverified:
      "Zu mindestens einem Treffer nennt TANSS keine Firmenzuordnung. Ob er wirklich zu " +
      "dieser Firma gehört, ist hier nicht feststellbar — bitte vor dem Anlegen prüfen.",
    remitterUnknown: "Der Absender ist in TANSS nicht bekannt.",
    remitterCreate: "Als Kontakt anlegen",
    remitterFirstName: "Vorname",
    remitterLastName: "Nachname",
    remitterMail: "E-Mail",
    remitterPhone: "Telefon",
    title: "Titel",
    textFold: "Titel und Text",
    content: "Beschreibung",
    quotedShow: "Zitierten Verlauf anzeigen",
    quotedHide: "Zitierten Verlauf ausblenden",
    type: "Tickettyp",
    status: "Status",
    priority: "Priorität",
    priorityDefault: "— Vorgabe von TANSS —",
    assignee: "Zuweisung",
    internal: "Intern",
    attachMail: "E-Mail am Ticket ablegen",
    similarHeading: "Ähnliche Tickets",
    similarHint: "Vielleicht gehört diese Mail an ein bestehendes Ticket.",
    similarSwitch: "Stattdessen anhängen",
    optionsFailures: "Nicht abrufbar:",
    submit: "Ticket anlegen",
    submitting: "Ticket wird angelegt …",
    successHeading: "Ticket {id} angelegt.",
    successMailByCheck:
      "Die Nachricht hängt am Ticket — nachgesehen. TANSS hat den Upload zwar bestätigt, " +
      "aber ohne Ergebniszeile geantwortet; deshalb wurde am Ticket geprüft statt " +
      "geraten. Bitte nicht erneut anhängen, sie wäre sonst zweimal da.",
    successMailAttached: "Die E-Mail hängt am Ticket.",
    successMailSkipped: "Die E-Mail wurde auf Wunsch nicht abgelegt.",
    successMailFailed: "Das Ticket wurde angelegt, die E-Mail konnte aber nicht abgelegt werden.",
    retryMail: "Mail jetzt anhängen",
  },

  ticketAnhaengen: {
    heading: "E-Mail an ein Ticket hängen",
    searchPlaceholder: "Ticketnummer oder Suchbegriff …",
    scopeCompany: "Nur diese Firma",
    scopeAll: "Alle Firmen",
    includeDone: "Abgeschlossene einbeziehen",
    recent: "Zuletzt verwendet",
    results: "Treffer",
    alreadyAttached: "Diese Mail hängt bereits an diesem Ticket.",
    attach: "Anhängen",
    attaching: "Mail wird abgelegt …",
    successHeading: "E-Mail an Ticket {id} abgelegt.",
    forceLabel: "Trotzdem anhängen",
    internal: "Als interne Mail kennzeichnen",
  },

  termin: {
    heading: "TANSS-Einsatz zu diesem Termin",
    resolving: "Zugehöriger Einsatz wird gesucht …",
    sourceProperty: "über die am Termin gespeicherte Einsatznummer gefunden",
    sourceLink: "über die lokale Zuordnungstabelle gefunden",
    sourceSyncgroup: "über die Kalendersynchronisation (SYNC_GROUP) gefunden",
    sourceTimematch: "über Startzeit und Dauer gefunden — bitte prüfen",
    sourceNone: "Zu diesem Termin gibt es noch keinen Einsatz.",
    candidates: "Mehrere Einsätze kommen in Frage. Bitte den richtigen wählen.",
    company: "Firma",
    ticket: "Ticket",
    ticketNone: "kein Ticket",
    type: "Einsatzart",
    location: "Ort",
    locationOFFICE: "Büro",
    locationCUSTOMER: "Beim Kunden",
    locationREMOTE: "Fernwartung",
    internal: "Intern",
    text: "Notiz",
    save: "Einsatz speichern",
    create: "Einsatz anlegen",
    createDisabled:
      "Das Anlegen von Einsätzen übernimmt in dieser Installation die Kalendersynchronisation.",
    savedHeading: "Einsatz {id} gespeichert.",
    createdHeading: "Einsatz {id} angelegt.",
    legacyMarkerFound:
      "Aus dem alten Add-in wurden Firma, Ticket und Einsatzart übernommen. " +
      "Der Marker im Terminkörper bleibt unverändert stehen.",
  },

  diagnose: {
    heading: "Diagnose",
    intro:
      "Diese Seite zeigt, was der laufende Outlook-Host tatsächlich meldet — nicht, was die " +
      "Dokumentation verspricht. Sie ist die erste Anlaufstelle, wenn etwas nicht funktioniert.",
    host: "Host",
    platform: "Plattform",
    officeVersion: "Office-Version",
    requirementSets: "Gemeldete Anforderungspakete",
    accountType: "Postfachart",
    accountTypeEnterprise: "enterprise (lokales Exchange) — Mail wird rekonstruiert",
    accountTypeCloud: "Exchange Online — Mail kommt über Microsoft Graph",
    mailbox: "Postfach",
    sessionStored: "TANSS-Sitzung im Speicher",
    sessionRefresh: "Erneuerungstoken vorhanden",
    sessionExpires: "Zugriffstoken gültig bis",
    sessionExpired: "abgelaufen",
    authMethod: "Benutzter Anmeldeweg",
    authClientId: "Entra-Anwendungs-Id",
    authResource: "Eigener Ressourcen-Bezeichner",
    authAccount: "Angemeldetes Konto",
    authLastError: "Letzter Anmeldefehler",
    bootSection: "Start",
    bootOffice: "Office-Bibliothek geladen",
    bootReady: "Antwort von Outlook",
    bootHostInfo: "Kennung des Hosts (_host_Info)",
    bootHostInfoMissing: "nicht übergeben",
    bootHostInfoInHash: "im Anker statt in der Adresse — office.js liest dort nicht",
    bootCsp: "Von der Inhaltsrichtlinie blockiert",
    bootCspNone: "nichts",
    bootCspLate: "erst spät zugehört — der Anfang fehlt",
    configuration: "Konfiguration",
    configApi: "TANSS-Adresse",
    configInstance: "Kennzeichen der Installation",
    configEntra: "Anwendungs-Id für Microsoft",
    configSource: "Herkunft der Einstellungen",
    service: "Dienst",
    serviceVersion: "Dienstfassung",
    employee: "TANSS-Mitarbeiter",
    limits: "Grenzen",
    maxEml: "Größte zulässige Mail",
    features: "Funktionen",
    featureAppointments: "Terminpflege",
    featureCreateSupport: "Einsatz anlegen",
    featureGraphMail: "Mailabruf über Graph",
    warnings: "Hinweise des Dienstes",
    copy: "Als Text kopieren",
    copied: "In die Zwischenablage kopiert.",
    copyFailed: "Kopieren hat nicht geklappt. Bitte den Text von Hand markieren.",
    notAvailable: "nicht verfügbar",
  },

  /**
   * Fehlertexte. Die Schluessel bis `INTERNAL` sind der geschlossene Code-Bestand des
   * Dienstes; sie stehen hier nur als Rueckfall, falls kein Koerper ankommt. Die Schluessel
   * mit Praefix `CLIENT_` entstehen ausschliesslich im Taskpane und erreichen den Dienst nie.
   */
  errors: {
    UNAUTHENTICATED:
      "Die Anmeldung ist nicht (mehr) gültig. Bitte das Fenster schließen und erneut öffnen.",
    IDENTITY_UNKNOWN:
      "Zu diesem Microsoft-Konto ist kein TANSS-Mitarbeiter hinterlegt. " +
      "Bitte den Administrator bitten, die Zuordnung einzutragen.",
    IDENTITY_AMBIGUOUS:
      "Zu diesem Microsoft-Konto passen mehrere TANSS-Mitarbeiter. " +
      "Es wird bewusst nicht geraten — der Administrator muss die Zuordnung festlegen.",
    FORBIDDEN_NO_COMPANY_ACCESS:
      "Für diese Firma fehlt in TANSS die Berechtigung. Der Vorgang wurde nicht ausgeführt.",
    REMITTER_REQUIRED: "Für diese Firma verlangt TANSS einen Melder.",
    TITLE_REQUIRED: "Ohne Titel kann kein Ticket angelegt werden.",
    MAIL_TOO_LARGE: "Die E-Mail ist größer als die zugelassene Höchstgröße.",
    MAIL_UNPARSABLE:
      "Die E-Mail ließ sich nicht auswerten und wurde deshalb nicht abgelegt — " +
      "sonst wäre sie in TANSS als gewöhnliches Dokument gelandet.",
    MAIL_ALREADY_ATTACHED: "Diese E-Mail hängt bereits an diesem Ticket.",
    MAIL_EMPTY:
      "Die Nachricht kam bei TANSS ohne Inhalt an und wurde nicht abgelegt. Das passiert " +
      "bei einer Mail, die Outlook noch nicht vollständig heruntergeladen hat — bitte " +
      "die Nachricht einmal öffnen und den Vorgang wiederholen.",
    MAIL_STORED_AS_DOCUMENT:
      "TANSS konnte die E-Mail nicht als Mail ablegen und hat sie als Dokument gespeichert. " +
      "In der Mailansicht des Tickets taucht sie deshalb nicht auf.",
    TICKET_CREATED_MAIL_FAILED:
      "Das Ticket wurde angelegt, die E-Mail konnte aber nicht abgelegt werden.",
    APPOINTMENT_NOT_SAVED:
      "Der Termin muss erst gespeichert werden, bevor er mit TANSS verknüpft werden kann.",
    APPOINTMENT_READONLY:
      "Dieser Termin ist in TANSS bereits zu einer Leistung geworden und wird nicht mehr geändert.",
    SUPPORT_AMBIGUOUS:
      "Zu diesem Termin passen mehrere Einsätze. Bitte den richtigen auswählen.",
    TANSS_UNAVAILABLE: "TANSS ist gerade nicht erreichbar. Der Vorgang wurde nicht ausgeführt.",
    TANSS_LICENSE_ERROR:
      "TANSS meldet ein Lizenzproblem. Das ist ein Fall für den Administrator, nicht für einen " +
      "zweiten Versuch.",
    CONFLICT:
      "Dieser Vorgang wurde bereits mit anderem Inhalt abgeschickt. " +
      "Bitte das Fenster schließen, neu öffnen und noch einmal beginnen.",
    RATE_LIMITED: "Zu viele Vorgänge in kurzer Zeit. Bitte einen Moment warten.",
    INTERNAL:
      "Im Dienst ist ein unerwarteter Fehler aufgetreten. Der Vorgang wurde nicht ausgeführt.",

    CLIENT_NETWORK:
      "Der TANSS-Dienst ist von Outlook aus nicht erreichbar. " +
      "Bitte die Netzverbindung prüfen und den Vorgang wiederholen.",
    CLIENT_BAD_RESPONSE:
      "Der Dienst hat geantwortet, aber nicht in der erwarteten Form. " +
      "Meist steckt ein Proxy oder eine Anmeldeseite dazwischen.",
    CLIENT_AUTH_UNAVAILABLE:
      "Dieser Outlook-Host unterstützt die Microsoft-Anmeldung nicht, die dieses Add-in braucht. " +
      "Einzelheiten stehen auf der Diagnoseseite.",
    CLIENT_AUTH_INTERACTION_REQUIRED:
      "Die Anmeldung braucht eine Rückfrage. Bitte das Anmeldefenster bestätigen.",
    CLIENT_AUTH_CONSENT_REQUIRED:
      "Für dieses Add-in fehlt die Zustimmung des Administrators. " +
      "Ohne sie erreicht es weder TANSS noch das Postfach.",
    CLIENT_AUTH_POPUP_BLOCKED:
      "Das Anmeldefenster wurde blockiert. Bitte es zulassen und den Vorgang wiederholen.",
    CLIENT_AUTH_FAILED:
      "Die Anmeldung am Microsoft-Konto ist fehlgeschlagen. " +
      "Einzelheiten stehen auf der Diagnoseseite.",
    CLIENT_GRAPH_FAILED:
      "Die E-Mail ließ sich nicht über Microsoft Graph holen. " +
      "Ohne sie kann nichts an ein Ticket gehängt werden.",
    CLIENT_MIME_FAILED:
      "Die E-Mail ließ sich in Outlook nicht zusammensetzen. " +
      "Der Vorgang wurde abgebrochen, damit nichts Halbes in TANSS landet.",
    CLIENT_ITEM_MISSING:
      "Outlook hat kein geöffnetes Element gemeldet. Bitte die E-Mail erneut öffnen.",
    CLIENT_OFFICE_FAILED:
      "Outlook hat eine Anfrage des Add-ins abgewiesen. Einzelheiten stehen auf der Diagnoseseite.",
    CLIENT_TOO_LARGE:
      "Die E-Mail ist größer als der Dienst annimmt. Sie wurde gar nicht erst übertragen.",
    CLIENT_UNEXPECTED:
      "Etwas ist schiefgegangen, womit dieses Add-in nicht gerechnet hat. " +
      "Der Vorgang wurde nicht ausgeführt.",
  },
});

/**
 * Text zu einem Punktpfad, z. B. `t("app.charCount", {n: 12, max: 100})`.
 *
 * Ein unbekannter Schluessel liefert den Schluessel selbst statt `undefined`: eine fehlende
 * Uebersetzung soll als Schoenheitsfehler auffallen und nicht als "undefined" in einem
 * Formular oder als Absturz beim Ersetzen.
 */
export function t(key, vars = null) {
  let node = T;
  for (const part of String(key).split(".")) {
    if (node === null || typeof node !== "object" || !(part in node)) return String(key);
    node = node[part];
  }
  if (typeof node !== "string") return String(key);
  if (!vars) return node;
  return node.replace(/\{(\w+)\}/g, (match, name) =>
    Object.prototype.hasOwnProperty.call(vars, name) ? String(vars[name]) : match,
  );
}

/**
 * Anzeigetext zu einem Fehler.
 *
 * Vorrang hat immer die Meldung des Dienstes - sie ist ein fertiger deutscher Satz und kennt
 * den konkreten Fall (welche Firma, welches Ticket). Die lokale Tabelle greift nur, wenn gar
 * kein Koerper ankam. Fehlt beides, steht dort ein ehrlicher Allgemeinplatz: ein leeres
 * Fehlerband, vor dem der Benutzer ratlos sitzt, ist das schlechtere Ergebnis.
 */
export function errorText(error) {
  if (!error) return T.errors.CLIENT_UNEXPECTED;
  const code = typeof error.code === "string" ? error.code : "";
  const message = typeof error.message === "string" ? error.message.trim() : "";
  // Ein `OfficeError` und ein `AuthError` tragen ihren CODE als Nachricht (`super(code)`).
  // Ohne diese Ausnahme stuende im Pane der blanke Bezeichner - "CLIENT_ITEM_MISSING" -
  // statt des deutschen Satzes, den diese Datei laut ihrem Kopf niemals umgehen lassen soll.
  if (message !== "" && message !== code) return message;
  if (code && Object.prototype.hasOwnProperty.call(T.errors, code)) return T.errors[code];
  return T.errors.CLIENT_UNEXPECTED;
}

/** Groessenangabe fuer Menschen. Eingabe in Bytes; drei Stufen reichen fuer Mails. */
export function formatBytes(bytes) {
  const n = Number(bytes);
  if (!Number.isFinite(n) || n < 0) return T.app.none;
  if (n < 1024) return t("app.bytes", { n });
  if (n < 1024 * 1024) return t("app.kilobytes", { n: Math.round(n / 1024) });
  return t("app.megabytes", { n: (n / (1024 * 1024)).toFixed(1).replace(".", ",") });
}

export default T;
