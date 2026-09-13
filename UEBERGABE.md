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
erzeugt. Einen Dienst gibt es nicht: Eine frühere Fassung, die einen Python-Dienst
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
  vendor/                 msal-browser, feste Fassung, Pruefsumme daneben
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
| TANSS-API ohne Token | HTTP 403 mit leerem Körper — der Client behandelt das als Fehler, nie als leere Trefferliste |
| Zertifikat einer nur intern auflösenden Instanz | Let's Encrypt, öffentlich vertrauenswürdig — Arbeitsplätze brauchen keine Zusatzmaßnahme |
| `X-Frame-Options` auf statischen Pfaden eines gewöhnlichen Apache | Wird gesendet — **muss** für die Ablage entfernt werden, sonst bleibt das Pane weiß |
| GitHub Pages | Sendet **kein** `X-Frame-Options`; setzt dafür gar keine Kopfzeilen, also kein `frame-ancestors`, und `Cache-Control` fest auf 10 Minuten |
| Parameter und Fragment in der Taskpane-Adresse | Überstehen beide; Office hängt sein `_host_Info` davor an und ersetzt nichts |
| Die Wächter über die Dateien | Schlagen nachweislich an: drei absichtlich eingebaute Verstöße, drei Fehlschläge |

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

**Das Unternehmenszeichen taugt nicht als Menübandsymbol.** Bei 16 px zerfällt die
Punktmarke zu Sprenkeln, bei 32 px ist sie grenzwertig. Nachgemessen mit gerenderten
Probestufen. Im Menüband steht deshalb das „T", das die Funktion benennt; das Zeichen
steht im Kopf des Panes.

---

## Offene Punkte

1. **Der Lauf in echtem Outlook steht aus.** Keine Maske ist je dort geöffnet worden;
   geprüft ist alles gegen Attrappen. Das ist der nächste Schritt und der wichtigste.
2. **Kein Git-Remote.** Ohne eines lässt sich nichts veröffentlichen. Danach einmalig
   *Einstellungen → Pages → Source: GitHub Actions*.
3. **Drei unbelegte Annahmen über TANSS** — sie betreffen den Server, nicht den
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
   der Dienstfassung entfallen. Der **natürliche Ort dafür ist der Prüflauf im
   Generator** — er spricht ohnehin schon mit TANSS.
4. **Der Generator hat keine eigenen Tests.** Seine Prüffunktion läuft bei jeder
   Erzeugung und verweigert ein schadhaftes Manifest, aber sie ist selbst ungeprüft. Ein
   kleines Testprojekt daneben wäre die naheliegende Ergänzung.
5. **Auf GitHub Pages fehlt `frame-ancestors`.** Der Hoster setzt keine Kopfzeilen. Das
   schützt nicht vor Tokendiebstahl — dafür bräuchte es einen Skriptfehler auf dem
   Ursprung selbst —, sondern nur vor Klickfallen. Wer das anders bewertet, nimmt die
   vhost-Vorlage unter `deploy/`.
6. **Ob eine Entra-Registrierung eingerichtet wird.** Ohne sie wird die Nachricht
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
