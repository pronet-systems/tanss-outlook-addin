# Übergabe

Arbeitsdokument für alle, die an diesem Repo weiterarbeiten. Es beschreibt den Stand,
die getroffenen Entscheidungen und — wichtiger — die Wege, die bereits gemessen und
verworfen wurden, damit niemand sie ein zweites Mal geht.

Kein Handbuch. Das Handbuch ist die `README.md`; sie ist für Betreiber geschrieben,
nicht für Entwickler.

Es enthält bewusst auch die Fehlschläge: Ein Weg, der einmal gemessen und verworfen
wurde, kostet beim zweiten Mal genauso viel Zeit wie beim ersten.

---

## Inhalt

- [Stand in einem Absatz](#stand-in-einem-absatz)
- [Wie es gebaut ist](#wie-es-gebaut-ist)
- [Was den Zuschnitt trägt](#was-den-zuschnitt-trägt)
- [Gemessen und belegt](#gemessen-und-belegt)
- [Gemessen und verworfen](#gemessen-und-verworfen)
- [Offene Punkte](#offene-punkte)
- [Regeln für die Arbeit an diesem Repo](#regeln-für-die-arbeit-an-diesem-repo)
- [Prüfbefehle](#prüfbefehle)

---

## Stand in einem Absatz

Das Add-in besteht aus statischen Dateien und einem Werkzeug, das daraus ein Manifest
erzeugt. Einen Dienst gibt es nicht: Eine frühere Version, die einen Python-Dienst
zwischen Outlook und TANSS legte, wurde vor der Veröffentlichung verworfen — ihre
Erfahrungen stecken in den Abschnitten unten. 141 Tests laufen durch, der Generator
baut, und das erzeugte Manifest nimmt Microsofts Prüfdienst an. Was fehlt, ist der Lauf
gegen ein echtes Postfach — **keine Maske ist je in Outlook geöffnet worden.**

---

## Wie es gebaut ist

```
taskpane/                 wird unveraendert ausgeliefert, rund 600 KB
  index.html              Inhaltsrichtlinie als <meta>; connect-src bewusst offen
  config.example.json     Vorlage; die echte config.json liegt nur auf dem Server
  js/
    api.js                VERTEILER ueber oertliche Vorgaenge
    runtime.js            Verdrahtung: Konfiguration, Sitzung, Client, Speicher
    config.js             liest config.json; entscheidet, welche Instanz gilt
    store.js              Vorlieben, zuletzt benutzte Tickets, Dublettenvermerke
    text.js               Betreff -> Tickettitel
    auth.js               NUR das Graph-Token fuer die Originalnachricht
    services/             die vierzehn fachlichen Vorgaenge
    tanss/                client, session, repository, models, uid, errors
    pages/                die fuenf Masken
  vendor/                 msal-browser, feste Version, Pruefsumme daneben
tools/ManifestGenerator/  WPF, erzeugt das Manifest und prueft die Angaben
deploy/                   vhost-Vorlage und die Konfiguration der gemeinsamen Ablage
tests/js/                 141 Faelle, eingebauter Testlaeufer von Node
.github/workflows/        CI und die Veroeffentlichung auf GitHub Pages
```

Es gibt **keine Abhängigkeit und keinen Bauschritt.** `package.json` erklärt die Dateien
zu ES-Modulen, damit der Testläufer genau die Dateien laden kann, die auch der Browser
lädt — mehr nicht, und ein Test misst das nach.

---

## Was den Zuschnitt trägt

Diese Punkte sind leicht zu übersehen und teuer zu verlieren.

**`api.js` ist ein Verteiler, keine `fetch`-Hülle.** Die Seiten rufen weiterhin
`api.get("api/search/tickets", {query})` und lesen `result.data.items`; dahinter liegen
örtliche Vorgänge. Diese Entscheidung hat beim Umbau rund 2.800 Zeilen Seitenlogik
unangetastet gelassen. Wer sie rückgängig macht, fasst alle fünf Masken an.

**Der Handelnde kommt ausschließlich aus der Sitzung.** Der TANSS-Client weist einen
übergebenen `loggedInUserId` als Programmfehler ab. Es gibt keinen zweiten Weg.

**`/api/v1/**` trägt ihn immer, `/api/tanss.x/v1/**` niemals.** Durchgesetzt im Client,
nicht verabredet. Mit dem Parameter nimmt der Server auf der zweiten Fläche andere
Zweige als die erprobten.

**Zwei Felder mit Löschwirkung sind gesperrt.** Die Teilnehmerliste eines Einsatzes geht
nie hinaus; die Zusatzinformationen nur als echte **Obermenge** des im selben Vorgang
gelesenen Satzes. Beides prüft eine Funktion unmittelbar vor dem Absenden.

**Die Terminkennung wird kanonisiert und nach der Regel des Servers verglichen.** Outlook
bettet die eigentliche Kennung in eine längere Kette ein; wer die Kette selbst
vergleicht, findet eine bestehende Kopplung nicht wieder. Und wer strenger prüft als der
Server, übersieht Kopplungen, die es gibt.

**Ein Erfolg ohne Beleg ist kein Erfolg.** Antwortet TANSS auf eine Anlage mit Erfolg,
aber ohne Nummer, wird das als Fehlschlag gemeldet. Nach dem Ablegen einer Nachricht
wird nachgesehen, ob daraus eine Mail oder still ein Dokument wurde.

**Zwei Zwischenspeicher.** Technikerliste 15 Minuten, Feldregeln 10 Minuten je Firma.
Das heißt auch: Ein Ausfall bleibt genauso lange unbemerkt. Zwei Testfälle halten beide
Seiten der Abwägung fest.

**Die Anmeldeseite nennt die Instanz, gegen die sie sich anmeldet.** Bei einer gemeinsam
genutzten Ablage ist die Adresse das Einzige, was die Instanzen unterscheidet — und die
nächste Eingabe ist ein Kennwort. Zugelassen ist nur `https`, optional nur eine Adresse
aus `allowedApiOrigins`.

---

## Gemessen und belegt

Nichts davon ist Annahme.

| Was | Ergebnis |
|---|---|
| Erzeugtes Manifest | Microsofts Prüfdienst: „The manifest is valid." Outlook Windows, Mac, Web |
| Fremder Ursprung an der TANSS-API | Vorabfrage HTTP 200, `Access-Control-Allow-Origin: *`, `apiToken` erlaubt |
| Der Erneuerungskopf `refreshToken` an derselben Vorabfrage | **Nicht** erlaubt — siehe unten. Anmeldung und Fachaufrufe gehen, die Erneuerung nicht |
| TANSS-API ohne Token | HTTP 403 mit leerem Körper — der Client behandelt das als Fehler, nie als leere Trefferliste |
| Zertifikat einer nur intern auflösenden Instanz | Let's Encrypt, öffentlich vertrauenswürdig — Arbeitsplätze brauchen keine Zusatzmaßnahme |
| `X-Frame-Options` auf statischen Pfaden eines gewöhnlichen Apache | Wird gesendet — **muss** für die Ablage entfernt werden, sonst bleibt das Pane weiß |
| GitHub Pages | Sendet **kein** `X-Frame-Options`; setzt dafür gar keine Kopfzeilen, also kein `frame-ancestors`, und `Cache-Control` fest auf 10 Minuten |
| Parameter und Fragment in der Taskpane-Adresse | Überstehen beide; Office hängt sein `_host_Info` davor an und ersetzt nichts |
| Die Wächter über die Dateien | Schlagen nachweislich an: drei absichtlich eingebaute Verstöße, drei Fehlschläge |
| `_host_Info` bei Adresse mit Parametern **und** Anker | Kommt im Abfrageteil an: `Outlook$Win32$16.02$de-DE$$$$0` — der Anker schadet nicht |
| Skript-Ursprünge, die office.js wirklich braucht | Zwei: der eigene **und** `ajax.aspnetcdn.com` — siehe unten |
| Welche TANSS-Flächen das Anmeldetoken eines Technikers erreicht | Nur `/api/v1/**`, einschließlich `/api/v1/admin/**`. **Nicht** `/api/tanss.x/v1/**` und **nicht** `/api/erp/v1/**` |
| Ticketpriorität | Ganzzahl 1–9 am Ticket, Feldname `priority`. Keine Tabelle, keine Namen — TANSS beschriftet mit der Ziffer. `0` beim Speichern ⇒ der Server setzt seine konfigurierte Vorgabe |
| Abteilungszugehörigkeit eines Technikers | Steht an der **Abteilung** (`employeeIds`), nicht am Mitarbeiter — dort führt TANSS nur die primäre |

### Die Erneuerung braucht eine Zeile auf dem TANSS-Server

Gemessen an `tanss.pronet-systems.de` (16.09.2026), Vorabfrage aus dem Ursprung
`https://pronet-systems.github.io`:

```
Access-Control-Allow-Headers: content-type,cache-control,x-requested-with,
                              x-xsrf-token,apiToken,user,password,
                              contractWorkflowToken,ticketWorkflowToken
```

`apiToken` steht darin, **`refreshToken` nicht** — und genau dieser Kopf ist der einzige
Weg, mit dem das Pane sein Zugriffstoken erneuert (`taskpane/js/tanss/session.js`).
Die Folge ist nicht sporadisch, sondern die Uhr: Anmeldung und alle Fachaufrufe gehen,
**vier Stunden** später läuft das Zugriffstoken ab, und die Erneuerung scheitert schon
in der Vorabfrage des Browsers.

Zu reparieren ist das **auf dem TANSS-Host**, nicht im Pane — im vhost, der `/backend`
ausliefert:

```apache
Header always set Access-Control-Allow-Headers "content-type,cache-control,x-requested-with,x-xsrf-token,apiToken,refreshToken,user,password,contractWorkflowToken,ticketWorkflowToken"
```

Danach `apachectl configtest`, `systemctl reload apache2`, und nachmessen:

```bash
curl -si -X OPTIONS "https://<tanss-host>/backend/api/v1/employees/ownState" \
  -H "Origin: https://<ursprung-des-panes>" \
  -H "Access-Control-Request-Method: GET" \
  -H "Access-Control-Request-Headers: refreshtoken" | grep -i allow-headers
```

Offen und beim Hersteller zu klären: ob TANSS diese Liste selbst konfigurierbar führt —
dann gehört der Eintrag dorthin statt in eine Apache-Zeile, die beim nächsten Umzug
vergessen wird. Ebenfalls ungeprüft, weil dafür ein gültiges Erneuerungstoken nötig
wäre: ob das Backend den Kopf hinter der Vorabfrage überhaupt auswertet.

Das Pane hält diesen Zustand seit dem Fund aus: Eine gescheiterte Erneuerung endet nicht
mehr in einer Fehlermeldung mit dem Knopf „Erneut versuchen" — der im CORS-Fall nie
wirken kann —, sondern führt auf dieselbe Anmeldeseite, auf der auch ein abgewiesenes
Token landet. Ein abgelaufenes Zugriffstoken, das sich nicht ersetzen lässt, heißt für
den Techniker dasselbe wie gar keines: Er kommt nur über eine Anmeldung weiter.

Die gespeicherte Sitzung wird dabei **nicht** verworfen — der Unterschied zum Abmelden,
und er ist beabsichtigt: War TANSS nur kurz nicht erreichbar, genügt auf dieser Seite
„Erneut versuchen", und ein Netzaussetzer kostet kein Kennwort.

Das ist ein Notausgang und kein Ersatz für die Zeile oben — ohne sie tippt jeder
Techniker alle vier Stunden sein Kennwort.

---

### Die Richtung der Ticketpriorität ist unbekannt — und bleibt es

Die Skala geht von 1 bis 9. Ob **1 oder 9 „hoch" bedeutet, geht aus nichts hervor**: Im
gesamten TANSS-Bestand wird die Priorität nirgends größer/kleiner verglichen, sondern nur
über Mengenzugehörigkeit; es gibt keinen Schwellenwert und keine Sortierung. Die einzige
Richtungsangabe der API-Dokumentation — *„priority of the callback from 1 (lowest) to 9
(highest)"* — gehört dem **Rückruf**, nicht dem Ticket.

Deshalb zeigt das Pane Ziffer und Sterne, so wie TANSS es selbst tut, und **kein einziges
Wort**. Wer keine Richtung behauptet, kann sie nicht verkehrt herum behaupten — und eine
verkehrt herum eingebaute Skala setzte jedes dringende Ticket auf die falsche Stufe, ohne
dass es auffiele. Ein Wächter hält Wörter wie „hoch" und „dringend" aus der Liste heraus.

Wer die Richtung belegen will, liest sie an einem echten Bestand ab — nicht an dieser
Dokumentation.

---

### Die Rolle entscheidet, nicht nur die Route

TANSS bindet jede Schnittstellenfläche an eine Rolle, und ein Anmeldetoken trägt **genau
eine**. Eine gewöhnliche Anmeldung ergibt die Rolle des Fachzugriffs:

| Fläche | Verlangte Rolle | Für uns erreichbar |
|---|---|---|
| `/api/v1/**` | Fachzugriff | ja |
| `/api/v1/admin/**` | Fachzugriff — „admin" ist ein Namensteil, keine Rollenforderung | ja |
| `/api/erp/v1/**` | ERP bzw. CENTRON | **nein** |
| `/api/tanss.x/v1/**` | Integrationsrolle | **nein** |

Daraus folgt eine Regel, die zweimal Tage gekostet hat: **Eine Route, die es gibt und die
sogar dokumentiert ist, kann für diesen Client trotzdem unerreichbar sein.** Wer eine
Route auswählt, prüft beides — ob es sie gibt *und* unter welcher Rolle sie liegt.

Zwei Aufrufe hingen daran. Die Technikerliste lief über die Integrationsfläche, weil jene
zusätzlich die Adresse liefert — die aber niemand braucht; sie läuft jetzt über
`/api/v1/employees/technicians`. Und die Tickettypen wurden kurzzeitig auf die
dokumentierte ERP-Route gelegt, die ein Techniker nie hätte aufrufen können; sie liegen
auf `/api/v1/admin/ticketTypes`. Ein Wächter hält beide Flächen jetzt aus den
Aufrufstellen heraus.

---

### Der Stillstand beim ersten Lauf in echtem Outlook

Das Pane lud sichtbar, `Office.onReady()` meldete sich aber nie, und alle Host-Felder
blieben leer. Ursache, vom Pane selbst gemeldet:

```
script-src-elem → https://ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js
```

office.js lädt für den Win32-Host **MicrosoftAjax von einem zweiten
Microsoft-Auslieferungsnetz** nach. Ob es das tut, hängt an einer Umschaltung im Dienst
(`EnableOfficeJsCDNForMsAjaxJs`): Ist sie an, kommt dieselbe Datei vom Ursprung von
office.js und der zweite Eintrag fällt niemandem auf. Ist sie aus, wird die Datei von
einer strengen Inhaltsrichtlinie abgewiesen — und dann kommt die Initialisierung nicht zum
Ende. Das Pane bleibt **stumm** stehen: Der Browser meldet den Verstoß nur seiner eigenen
Konsole, und die gibt es im eingebetteten Browser von Outlook nicht.

Die restliche Ladekette wurde nachgesehen und liegt vollständig auf dem Ursprung von
office.js: `o15apptofilemappingtable.js`, `outlook-15.02.js`, `ariatelemetry/`. Nicht
eingetragen ist `alcdn.msauth.net` — er steht in der Vertrauensliste von office.js, gehört
aber zum alten Office-Anmeldeweg, den dieses Pane nicht benutzt.

Der Telemetrierahmen von office.js
(`telemetryservice.firstpartyapps.oaspapps.com`) wird von `frame-src` **absichtlich**
abgewiesen. Er wird drei Sekunden nach dem Start nachgeladen, hält nichts auf und trägt
nichts bei. In der Diagnose taucht er als blockiert auf und ist dort kein Befund.

**Lehre für jede weitere Arbeit:** Eine strenge Inhaltsrichtlinie und ein Add-in-Host, der
Bibliotheken von wechselnden Ursprüngen nachlädt, vertragen sich nur, wenn das Pane selbst
meldet, was abgewiesen wurde. Genau dafür gibt es `taskpane/js/boot-early.js` — ein
klassisches Skript **vor** office.js, weil ein Modul zurückgestellt liefe und die
Meldungen verpasste. Zwei Wächter halten es an seinem Platz.

---

## Gemessen und verworfen

**Die lokale Registrierung eines Outlook-Add-ins trägt nicht.** Der Registry-Schlüssel
`HKCU\Software\Microsoft\Office\16.0\WEF\Developer` ist ein *Vermerk* über eine erfolgte
Installation, kein Installationsweg. Gemessen: Manifest gültig, Eintrag in der
dokumentierten Form gesetzt, NTFS-Stundensperre durch Zurücksetzen der Zugriffszeit
ausgehebelt, Gegenprobe am eigenen Lesevorgang bestanden — klassisches Outlook hat die
Manifestdatei in sechs Minuten mit verbundenem Postfach **kein einziges Mal geöffnet**.
Die Dokumentation bestätigt es: Bei einem Outlook-Add-in wird das Manifest im
Exchange-Dienst registriert.

**Folge:** Die Registrierung läuft einmalig über das Admin Center und ist durch keine
lokale Maßnahme zu ersetzen.

**Der Anker in der Taskpane-Adresse ist nicht die Ursache von Startproblemen.** Die
Vermutung lag nahe: Die Manifestadressen enden auf `#/ticket/neu`, und office.js liest
`_host_Info` ausschließlich aus `window.location.search`. Gemessen ist sie trotzdem
falsch — die Kennung kam vollständig im Abfrageteil an. Eine laufende Fremdinstallation
bestätigt es: Sie arbeitet ebenfalls mit Hash-Routen.

**Das Unternehmenszeichen taugt nicht als Menübandsymbol.** Bei 16 px zerfällt die
Punktmarke zu Sprenkeln, bei 32 px ist sie grenzwertig. Nachgemessen mit gerenderten
Probestufen. Im Menüband steht deshalb das „T", das die Funktion benennt; das Zeichen
steht im Kopf des Panes.

---

## Offene Punkte

1. **Der Lauf in echtem Outlook ist begonnen, nicht abgeschlossen.** Das Menüband
   erscheint, das Pane lädt, der Start kommt durch (siehe den Befund zu MicrosoftAjax
   oben). Was noch nie dort gelaufen ist: Anmeldung an TANSS, Ticketanlage, Mailablage,
   Terminpflege. Alles darunter ist gegen Attrappen geprüft und sonst nirgends.
2. **Drei unbelegte Annahmen über TANSS** — sie betreffen den Server, nicht den
   Zuschnitt, und fallen alle sichtbar aus statt still:
   - ob `PUT /api/v1/supports/list` das Feld `fetchMetaInfos` annimmt und die
     Zusatzinformationen zurückgibt. Bleiben sie aus, beruht die Terminzuordnung allein
     auf dem Zeitfenster — und die Maske sagt das.
   - ob `GET /api/v1/tickets/` die Feldsteuerung im `meta.properties`-Block liefert.
     Tut sie es nicht, gilt keine Melderpflicht und keine Zuweisungspflicht: eine zu
     milde Maske, kein falscher Inhalt.
   - wo der Hinweis auf eine gekappte Suche steht. Gelesen werden zwei Behälter; steht
     er an einer dritten Stelle, ist die Folge „nichts gefunden" statt „Suche
     verfeinern".

   Ein Python-Werkzeug, das diese drei gegen eine laufende Instanz gemessen hat, ist mit
   der Dienstversion entfallen. Der **natürliche Ort dafür ist der Prüflauf im
   Generator** — er spricht ohnehin schon mit TANSS.
3. **Der Generator hat keine eigenen Tests.** Seine Prüffunktion läuft bei jeder
   Erzeugung und verweigert ein schadhaftes Manifest, aber sie ist selbst ungeprüft. Ein
   kleines Testprojekt daneben wäre die naheliegende Ergänzung.
4. **Auf GitHub Pages fehlt `frame-ancestors`.** Der Hoster setzt keine Kopfzeilen. Das
   schützt nicht vor Tokendiebstahl — dafür bräuchte es einen Skriptfehler auf dem
   Ursprung selbst —, sondern nur vor Klickfallen. Wer das anders bewertet, nimmt die
   vhost-Vorlage unter `deploy/`.
5. **Ob eine Entra-Registrierung eingerichtet wird.** Ohne sie wird die Nachricht
   rekonstruiert statt kopiert; Empfangskette und Signaturköpfe fehlen dann. Die
   Registrierung braucht kein Geheimnis, und der Generator prüft sie mit.

---

## Regeln für die Arbeit an diesem Repo

**Keine Herkunftsbezüge.** Weder in Quelltext, Kommentaren, Commits noch Dokumentation
steht, woher das Wissen über die TANSS-API stammt. Kommentare beschreiben **Verhalten**,
nicht Quellen. Ebenso keine Vergleiche mit einem Hersteller oder einem Vorbild.

**Kein Kundenname im Repo.** Es wird als MIT-Quelltext veröffentlicht. Adressen gehören
in die `config.json` auf dem Server und in das Manifest — beides Ablage, nicht Quelltext.
Ein Test lässt nur reservierte Beispieladressen zu.

**Keine Attribution im Git.** Commit-Nachrichten tragen keine `Co-Authored-By`-,
`Claude-Session`- oder „Generated with"-Zeile.

**Keine Abhängigkeit.** Weder im Taskpane noch im Generator. Jede wäre ein weiterer
Herausgeber von Code, der dort läuft, wo Betreff, Klartext und ein Token durchgehen.

**Sprache.** Dokumentation in Markdown mit Umlauten. Quelltextkommentare ohne Umlaute
(`ae`, `oe`, `ue`, `ss`) — durchgehend so, und das soll einheitlich bleiben.

**Tests.** Ein Test, der nach einer Änderung nichts mehr prüft, ist schlimmer als kein
Test. Schlägt ein Wächter an, wird er **begründet nachgezogen, nicht stillgelegt.** Und
wer einen neuen Wächter schreibt, weist einmal nach, dass er auch anschlägt — ein
Suchlauf über null Dateien ist grün und sagt nichts.

---

## Prüfbefehle

```bash
npm test                                            # 141 Faelle
cd tools/ManifestGenerator && dotnet build          # der Generator
```

Manifest gegen Microsofts Prüfdienst:

```bash
npx office-addin-manifest validate <pfad>/manifest.xml
```

Gegen eine laufende Instanz misst der Generator unter **Prüfen** — lesend, ohne
Anmeldung.

Stand: 13.09.2026.
