# TANSS Outlook Add-in

Ein Outlook-Add-in für **TANSS**. Es legt E-Mails samt Anhängen an Tickets ab, erzeugt
aus einer E-Mail ein Ticket und pflegt zu einem Outlook-Termin den zugehörigen
TANSS-Einsatz.

Es besteht aus **statischen Dateien** — HTML, CSS und JavaScript. Kein Dienst, kein
Port, kein systemd, keine Datenbank, kein Geheimnis im Ruhezustand. Sie legen ein
Verzeichnis auf einem Webserver ab und laden einmalig eine Manifestdatei in Ihr
Microsoft 365 Admin Center. Das Taskpane spricht danach direkt mit Ihrer TANSS-Instanz,
aus dem Browser des Technikers heraus.

> **Status:** Vor dem ersten Produktivbetrieb. Für die Inbetriebnahme empfiehlt sich ein
> Pilotbetrieb: das Add-in zunächst einer kleinen Gruppe zuweisen, einige Tage
> beobachten, dann ausrollen.

---

## Inhalt

- [Was das Add-in kann](#was-das-add-in-kann)
- [Wie es aufgebaut ist](#wie-es-aufgebaut-ist)
- [Was dieses Werkzeug nicht tut](#was-dieses-werkzeug-nicht-tut)
- [Voraussetzungen](#voraussetzungen)
- [Installation](#installation)
  - [1. Dateien ablegen](#1-dateien-ablegen)
  - [2. Konfiguration eintragen](#2-konfiguration-eintragen)
  - [3. Manifest erzeugen](#3-manifest-erzeugen)
  - [4. In Outlook bereitstellen](#4-in-outlook-bereitstellen)
- [Mehrere Kunden aus einer Ablage](#mehrere-kunden-aus-einer-ablage)
- [Konfiguration im Einzelnen](#konfiguration-im-einzelnen)
- [Die Originalnachricht](#die-originalnachricht)
- [Aktualisieren](#aktualisieren)
- [Wenn etwas nicht geht](#wenn-etwas-nicht-geht)
- [Entwicklung](#entwicklung)
- [Lizenz](#lizenz)

---

## Was das Add-in kann

Im Lesemodus einer E-Mail erscheint **eine** Schaltfläche „TANSS" mit einem
Aufklappmenü; darin stehen *Ticket erstellen* und *An Ticket anhängen*. Im Termin des
Organisators erscheint eine zweite Schaltfläche. Alle Einstiege öffnen dasselbe Taskpane
auf einer eigenen Seite.

> Warum ein Menü und nicht zwei Schaltflächen: Das Menüband einer Mailansicht ist voll.
> Bei zwei Schaltflächen klappt Outlook als erstes die Beschriftungen weg, und es stehen
> zwei fast gleiche Symbole nebeneinander, die niemand auseinanderhält. Im Menü steht
> jeder Vorgang ausgeschrieben da — und die beiden Einträge tragen eigene Symbole.

**E-Mail an ein bestehendes Ticket hängen.** Der häufigste Fall: Acht Nachrichten
gehören zu einem Vorgang, nicht zu acht Tickets. Das Taskpane sucht Tickets über Nummer
oder Titel, bietet zuletzt verwendete an und legt die geöffnete Nachricht mit allen
Anhängen als Mail am Ticket ab. Anschließend sieht es nach, was daraus geworden ist:
Konnte TANSS die Nachricht nicht als Mail lesen, legt es sie still als gewöhnliches
Dokument ab — in der Mailansicht des Tickets wäre sie dann unsichtbar, und genau das
meldet das Taskpane, statt Erfolg zu behaupten.

**Aus einer E-Mail ein Ticket anlegen.** Firma und Melder werden aus der Absenderadresse
aufgelöst und mit Herkunftsbegründung angezeigt; ist der Absender in TANSS unbekannt,
lässt er sich im selben Schritt als Kontakt anlegen. Betreff und Text werden vorbelegt,
der zitierte Vorgänger abgetrennt und eingeklappt angeboten. Ticketanlage und Ablage der
Mail laufen als **ein** Vorgang: Ein Doppelklick erzeugt kein zweites Ticket, und wenn
nur die Ablage scheitert, nennt die Erfolgsseite trotzdem die Ticketnummer und bietet
die Wiederholung an. Gibt es bereits ähnliche offene Tickets derselben Firma, weist ein
Band darauf hin und führt mit einem Klick in die Anhängen-Maske.

**Zum Outlook-Termin den TANSS-Einsatz pflegen.** Das Taskpane sucht den Einsatz, der zu
diesem Termin gehört, und lässt Firma, Einsatzart, Ticketbezug, Ort und die
Kennzeichnung *intern* daran ändern. Gefunden wird er über die am Termin hinterlegte
Kennung, über die in TANSS geführte Kopplung und zuletzt über einen Abgleich von Beginn
und Dauer. Mehrere Treffer ergeben eine **Auswahlliste, kein Raten**. Der sichtbare
Terminkörper wird nie verändert, die Teilnehmerliste nie angefasst.

**Diagnose im Taskpane.** Die Seite `#/diagnose` zeigt, was der laufende Outlook-Client
tatsächlich meldet: Host und Version, die verfügbaren Anforderungspakete, den Ort des
Postfachs und den angemeldeten TANSS-Mitarbeiter. Sie ist ohne Anmeldung erreichbar,
weil die häufigste Frage an sie lautet, warum die Anmeldung gerade nicht funktioniert.
Token, Betreffzeilen und Mailinhalte stehen dort nicht.

---

## Wie es aufgebaut ist

Es gibt **keine Gegenstelle zwischen Outlook und TANSS**. Das Taskpane läuft im Browser
des Technikers, meldet sich dort an TANSS an und spricht von dort aus direkt mit der
TANSS-API.

| Bestandteil | Wo | Inhalt |
|---|---|---|
| Taskpane | Ein Verzeichnis auf einem beliebigen Webserver | HTML, CSS, ES-Module, Symbole. Rund 600 KB. Wird unverändert ausgeliefert — kein Buildschritt, kein npm, kein Node auf dem Server |
| Konfiguration | `config.json` neben `index.html` | Adresse der TANSS-Instanz, Grenzen, Vorgaben. **Öffentlich lesbar und ohne Geheimnis** |
| Manifest | `manifest.xml`, einmalig ins Admin Center geladen | Sagt Outlook, wo das Taskpane liegt und wie die Schaltflächen heißen |
| Anmeldung | Im Browser des Technikers | Sein eigenes TANSS-Token, im lokalen Speicher seines Arbeitsplatzes |
| Vorbelegungen | Im Browser des Technikers | Zuletzt benutzter Tickettyp, zuletzt benutzte Tickets, Dublettenvermerke |

Der Webserver, der die Dateien ausliefert, **spricht nie mit TANSS**. Er weiß nicht
einmal, wohin sich das Taskpane verbindet. Daraus folgt unmittelbar, dass Ihre
TANSS-Instanz nicht öffentlich erreichbar sein muss: Erreichbar sein muss sie für die
**Arbeitsplätze der Techniker**, und die stehen in Ihrem Netz.

---

## Was dieses Werkzeug nicht tut

Jede dieser Zusagen hat einen Mechanismus, der sie trägt, nicht nur eine Absicht.

| Zusage | Wodurch sie getragen wird |
|---|---|
| **Keine Telemetrie an Dritte.** Keine Nutzungsdaten, Betreffzeilen, Mailtexte oder Fehlerberichte an fremde Stellen. | Im Taskpane gibt es keinen Fehlersammler und keinen fest verdrahteten fremden Host. Ein unerwarteter Clientfehler geht in die Konsole dieses Rechners und sonst nirgendwohin. Ein Prüflauf durchsucht `taskpane/` nach absoluten Adressen außerhalb der erlaubten Liste und schlägt bei jeder weiteren fehl. |
| **Keine Serverkomponente, die Postfächer erreichen könnte.** | Es gibt keine. Die Nachricht wandert vom Browser des Technikers direkt zu TANSS; kein Zwischensystem sieht sie, kein Dienst kann übernommen werden. |
| **Kein Geheimnis im Ruhezustand.** Kein Client-Secret, kein Dienst-Token, kein Zertifikatsschlüssel, der ablaufen oder abfließen könnte. | Der Techniker meldet sich selbst an und bekommt sein eigenes Token. Die `config.json` enthält ausschließlich Adressen und Vorgaben und darf öffentlich gelesen werden. |
| **Kein Anwendungstoken im Browser.** | Das Anmeldetoken eines Technikers trägt seinen Mitarbeiterbezug in der signierten Aussage und ist damit nicht fälschbar. Ein Anwendungstoken, das den Mitarbeiter als Parameter mitführt, gehört nie in einen Browser — es kommt in diesem Add-in nicht vor. |
| **Keine Löschungen in TANSS.** Weder Tickets noch Einsätze, Mails, Dokumente oder Kontakte werden entfernt. | Kein Codepfad ruft eine löschende TANSS-Route auf. Beim Speichern eines Einsatzes wird ein geschlossenes Feldobjekt gesendet: Die Teilnehmerliste kann gar nicht durchgereicht werden, und die Zusatzinformationen werden nur geschrieben, wenn der neue Satz eine **Obermenge** des im selben Vorgang gelesenen ist. Beides ist durch eine Prüfung unmittelbar vor dem Absenden abgesichert. |
| **Kein Verschieben von Terminen.** Das Add-in reichert einen Termin um TANSS-Felder an. | Beim Ändern eines Einsatzes gehen weder Beginn noch Dauer noch Mitarbeiter noch Planungsart mit. Was nicht gesendet wird, kann nichts überschreiben. |
| **Keine Mailinhalte im Ruhezustand.** | Es gibt keinen Ort, an dem sie liegen könnten. Der lokale Speicher des Arbeitsplatzes hält von einer abgelegten Nachricht nur einen Streuwert ihrer Kennung — genug für „schon einmal angehängt", zu wenig, um daraus einen Korrespondenzpartner zu lesen. |

**Was es dafür verlangt, und was das frühere Modell nicht verlangte:** Der Techniker
meldet sich mit seinen **TANSS-Zugangsdaten** im Taskpane an. Das ist der Preis dafür,
dass es keine Serverkomponente mehr gibt, die stellvertretend handeln könnte. Sein Token
liegt danach im lokalen Speicher seines Arbeitsplatzes, gilt vier Stunden und wird
selbsttätig erneuert; eine neue Anmeldung steht erst an, wenn auch das Erneuerungstoken
abgelaufen ist.

Daraus folgt eine Sorgfaltspflicht, die die Anmeldeseite sichtbar macht: Sie **zeigt
an, gegen welche TANSS-Instanz** sie sich anmeldet. Eine untergeschobene Adresse wäre
eine Anmeldemaske, die das Kennwort woanders hinschickt — deshalb nimmt das Taskpane nur
`https`-Adressen an und lässt sich auf eine Positivliste einschränken
(`allowedApiOrigins`).

Ebenfalls nicht enthalten: Mobilunterstützung, ein Verfassen-Modus für neue Nachrichten,
eine zweite Oberflächensprache, ein Offline-Zwischenspeicher und eine Warteschlange.

---

## Voraussetzungen

**Auf Serverseite** — nur für die Ablage der Dateien:

- Ein Webserver, der statische Dateien über **HTTPS** ausliefert. Es genügt jeder:
  Apache, nginx, ein Objektspeicher, GitHub Pages.
- Ein Zertifikat, dem die **Arbeitsplätze** vertrauen.
- Die Möglichkeit, für diesen Pfad **keine** `X-Frame-Options`-Kopfzeile zu senden.
  Sendet der Server sie, bleibt das Taskpane in Outlook weiß — siehe
  [Wenn etwas nicht geht](#wenn-etwas-nicht-geht).

**Auf TANSS-Seite:**

- TANSS **10.10** oder neuer, erreichbar von den Arbeitsplätzen der Techniker.
- Ein TANSS-Konto je Techniker. Weitere Rechte als die, die er ohnehin hat, braucht er
  nicht — das Add-in handelt als er.

**Auf Microsoft-Seite:**

- Ein Microsoft 365 Mandant mit Exchange Online.
- Administratorzugang zum Admin Center für die einmalige Bereitstellung.
- Optional eine Entra-Registrierung, siehe [Die
  Originalnachricht](#die-originalnachricht).

**Outlook:** klassisches Outlook für Windows, neues Outlook für Windows, Outlook im
Browser und Outlook für Mac. Mindestens Mailbox-Anforderungspaket 1.6.

---

## Installation

### 1. Dateien ablegen

Der gesamte Inhalt von `taskpane/` kommt in ein Verzeichnis, das über HTTPS erreichbar
ist. Nichts davon wird vorher gebaut oder übersetzt.

```bash
rsync -a --delete taskpane/ /var/www/tanss-outlook-addin/
```

Für einen eigenen vhost auf einem Apache liegt eine vollständig kommentierte Vorlage
bereit: `deploy/apache-addon-vhost.conf.example`. Sie enthält die beiden Einstellungen,
an denen es sonst scheitert — das Entfernen von `X-Frame-Options` und eine
Inhaltsrichtlinie mit `frame-ancestors` für die Office-Herkünfte.

#### Oder über GitHub Pages

Dieses Repository bringt einen Ablauf mit, der das Taskpane nach jedem Push auf `main`
veröffentlicht (`.github/workflows/pages.yml`). Einzurichten ist einmal:

**Einstellungen → Pages → Source: GitHub Actions.**

Danach liegt das Pane unter `https://<konto>.github.io/<repo>/`. Es wird nichts gebaut —
veröffentlicht werden genau die Dateien aus `taskpane/`, und vorher laufen beide
Testsuiten.

Zwei Unterschiede zum eigenen vhost, beide gemessen und beide verkraftbar:

| | GitHub Pages | eigener vhost |
|---|---|---|
| `X-Frame-Options` | sendet keine ✓ | muss entfernt werden |
| `frame-ancestors` | nicht setzbar — jede Seite dürfte das Pane einrahmen | setzbar |
| `Cache-Control` | fest 10 Minuten | frei wählbar |
| Aufwand | ein Push | DNS, Zertifikat, vhost |

Der Verlust von `frame-ancestors` schützt nicht vor Tokendiebstahl — dafür bräuchte es
einen Skriptfehler auf dem Ursprung selbst — sondern nur vor Klickfallen, und die
Vorgänge hier sind alle mehrschrittig. Wer das anders bewertet, nimmt den eigenen vhost;
die Vorlage bleibt dafür im Repository.

Auf einer so geteilten Ablage liegt **keine Kundenadresse**: Die Veröffentlichung legt
eine `config.json` daneben, die `apiBase` leer lässt. Welche Instanz gilt, entscheidet
dann das Manifest jedes Kunden — siehe [Mehrere Kunden aus einer
Ablage](#mehrere-kunden-aus-einer-ablage).

### 2. Konfiguration eintragen

```bash
cp /var/www/tanss-outlook-addin/config.example.json \
   /var/www/tanss-outlook-addin/config.json
```

Dann `config.json` bearbeiten. Zwei Angaben genügen für den Anfang:

```json
{
  "apiBase": "https://tanss.ihre-firma.de/backend",
  "frontendBase": "https://tanss.ihre-firma.de"
}
```

Die Datei darf fehlen, darf Unsinn enthalten und darf unbekannte Schlüssel tragen — das
Taskpane öffnet in jedem Fall und sagt auf der Diagnoseseite, was es davon übernommen
hat. Ein Tippfehler legt keine neue Einstellung an, sondern wird verworfen und benannt.

### 3. Manifest erzeugen

Das Manifest sagt Outlook, wo die Dateien liegen und gegen welche TANSS-Instanz
gearbeitet wird. Es wird **erzeugt**, nicht von Hand gepflegt.

```bash
cd tools/ManifestGenerator
dotnet run
```

Für eine Einzeldatei, die ohne installierte .NET-Laufzeit läuft:

```bash
dotnet publish -c Release -r win-x64 --self-contained true \
  -p:PublishSingleFile=true -p:EnableCompressionInSingleFile=true
```

Im Fenster sind zwei Felder Pflicht: die Ablage-Adresse der Dateien und die TANSS-API.
Die **Add-in-Kennung** darunter wird aus der TANSS-Adresse berechnet, nicht gewürfelt —
dieselbe Instanz ergibt immer dieselbe Kennung, auch in zwei Jahren und auf einem
anderen Rechner. Das ist wichtig: Für Outlook ist ein Add-in mit anderer Kennung ein
**anderes** Add-in, und die alte Installation bliebe bei jedem Benutzer daneben stehen.

Die Schaltfläche **Prüfen** misst die Angaben gegen die Wirklichkeit, lesend und ohne
Anmeldung:

| Bereich | Geprüft wird |
|---|---|
| Manifest | Struktur, Berechtigungsstufe, getrennte Herkunft von Pane und TANSS |
| Ablage | erreichbar, Symbole vorhanden, `X-Frame-Options`, `frame-ancestors` |
| TANSS | API antwortet unter dem angegebenen Pfad, fremder Ursprung zugelassen |
| Zertifikat | Ablaufdatum beider Adressen |
| Anmeldung | Entra-Anwendung vorhanden (braucht den Mandanten) |

Ein dritter Ausgang neben „in Ordnung" und „beanstandet" ist **„unklar"**: Eine Prüfung,
die keine Antwort bekommen hat, gilt nicht als bestanden. Die Rückadresse der
Entra-Registrierung steht dauerhaft auf „unklar" — der Anmeldedienst beanstandet eine
fehlende Rückadresse gegenüber einem Unangemeldeten nicht, sondern erst nach der
Anmeldung. Statt eines „OK", das nichts belegt, nennt der Prüflauf dort den genauen
Wert zum Nachsehen.

Alle Prüfungen kommen ohne Zugangsdaten aus und verändern nichts. Eine ist nicht rein
lesend: Die Prüfung der Entra-Anwendung fordert einen Anmeldecode an, um zu erfahren, ob
es die Anwendung gibt. Der Code wird nicht benutzt und verfällt; einlösen könnte ihn
ohnehin nur, wer gültige Zugangsdaten hat.

### 4. In Outlook bereitstellen

Microsoft 365 Admin Center → **Einstellungen** → **Integrierte Apps** → **Add-Ins** →
**Add-In bereitstellen** → **Benutzerdefiniertes Add-In** → **Manifestdatei hochladen**.

Die Datei wird **hochgeladen, nicht verlinkt**. Ihr Webserver muss für Microsoft also
nie erreichbar sein — nur für die Arbeitsplätze Ihrer Techniker.

Danach den Benutzern oder einer Gruppe zuweisen. Bis die Schaltflächen im Menüband
erscheinen, können 24 bis 72 Stunden vergehen; ein Neustart von Outlook beschleunigt es
oft.

---

## Mehrere Kunden aus einer Ablage

Weil der Hoster nie selbst mit TANSS spricht, kann **eine einzige Ablage beliebig viele
Instanzen bedienen** — auch die eines anderen Unternehmens, dessen TANSS nur in dessen
eigenem Netz erreichbar ist.

Dafür trägt der Generator die TANSS-Adresse in das Manifest ein. Jeder Kunde bekommt
sein eigenes Manifest und lädt es in sein eigenes Admin Center; auf dem Hoster liegt
nichts Kundenspezifisches.

Welche Instanz gilt, entscheidet sich in drei Stufen:

1. **Die Adresse im Manifest.** Sie stammt vom Administrator des Mandanten und ist damit
   verwaltet, nicht geraten.
2. **`apiBase` aus der `config.json`**, für eine Installation mit eigenem Hoster.
3. **Eine einmal selbst eingetragene Adresse**, falls weder das eine noch das andere
   vorliegt.

In allen drei Fällen gilt: nur `https`, und nur ein Ursprung aus `allowedApiOrigins`,
sofern diese Liste gefüllt ist.

---

## Konfiguration im Einzelnen

| Schlüssel | Bedeutung |
|---|---|
| `apiBase` | Adresse der TANSS-API, meist auf `/backend` endend. Leer lassen, wenn sie aus dem Manifest kommt. |
| `allowedApiOrigins` | Liste zugelassener Ursprünge. Leer: jede `https`-Adresse. |
| `frontendBase` | Die TANSS-Weboberfläche. Daraus entstehen die Verweise „in TANSS öffnen". Leer: keine Verweise. |
| `instance` | Kennzeichen dieser Installation. Es trennt die Terminkopplungen zweier Installationen auf demselben Arbeitsplatz. Leer: aus dem Hostnamen abgeleitet. |
| `entra.clientId` | Anwendungs-Id für die Microsoft-Anmeldung. Siehe unten. |
| `limits.maxEmlBytes` | Größte Nachricht, die abgelegt wird. Vorgabe 25 MiB. |
| `features.appointments` | Die Terminseite anbieten. |
| `features.createSupport` | Selbst Einsätze anlegen. **Nur** in einer Installation ohne Kalendersynchronisation einschalten — sonst entstehen zwei Einträge je Termin. |
| `features.graphMail` | Die Nachricht über Microsoft Graph holen. Braucht `entra.clientId`. |
| `defaults.*` | Vorbelegung von Einsatzart, Ort, Intern-Kennzeichen und Tickettyp. |
| `ticketTypes` | Die Tickettypen dieser Installation als `[{id, name}]`. Leer: Die Ticketmaske zeigt kein Typfeld. |

---

## Die Originalnachricht

Für die Ablage am Ticket gibt es zwei Wege, und sie unterscheiden sich in der Qualität
des Ergebnisses.

**Mit** einer Entra-Registrierung holt das Taskpane die Nachricht im Originalformat über
Microsoft Graph. Das ist eine Kopie: Empfangskette, Signaturköpfe und Referenzen bleiben
erhalten.

**Ohne** sie setzt das Taskpane die Nachricht aus Office-Bordmitteln zusammen. Das ist
eine **Rekonstruktion**: Der Inhalt und die Anhänge sind da, aber `Received`-Kette,
`DKIM-Signature` und `References` fehlen. Für Beweiszwecke ist eine so abgelegte Mail
schwächer als eine, die der Mailroboter importiert hat.

Die Registrierung braucht **kein Geheimnis**. Sie ist eine öffentliche Clientanwendung
mit der delegierten Berechtigung `Mail.Read` und einer Rückadresse unter der Plattform
*Einseitige Anwendung (SPA)*:

```
brk-multihub://<hostname-ihrer-ablage>
```

Nur die **Herkunft**, ohne Pfad. Liegt das Pane unter
`https://beispiel.github.io/tanss-outlook-addin`, lautet die Rückadresse
`brk-multihub://beispiel.github.io` — nicht die volle Adresse.

Die Anwendungs-Id tragen Sie im Generator ein; sie wandert in die Adressen des
Manifests. Im Manifest selbst steht **kein** Anmeldeblock: Das Pane holt sein Token über
die verschachtelte Anmeldung, und die verlangt dafür nichts im Manifest.

Ob es die Anwendung gibt, prüft der Generator mit; dafür braucht er im Feld *Mandant*
Ihre Microsoft-365-Domäne oder die Verzeichnis-ID. Die **Rückadresse lässt sich von
außen nicht prüfen** — Microsoft gibt sie Unangemeldeten nicht preis. Sie ist zugleich
die häufigste Fehlerquelle der ganzen Einrichtung, und der Prüflauf nennt deshalb den
genauen Wert, den Sie unter *Authentifizierung* nachsehen sollten.

---

## Aktualisieren

**Die Dateien** tauschen Sie jederzeit aus; beim nächsten Öffnen des Panes gilt die neue
Fassung. Ein Manifest muss dafür nicht angefasst werden.

```bash
rsync -a --delete taskpane/ /var/www/tanss-outlook-addin/   # config.json ausnehmen!
```

**Das Manifest** nur dann, wenn sich Beschriftung, Menüband, Adresse oder
Berechtigungsstufe ändern. Dann im Generator die **Version erhöhen**, neu erzeugen und
im Admin Center über **Update** einspielen.

> Bleibt die Version gleich, übergeht Microsoft die Aktualisierung **stillschweigend**,
> und die Benutzer behalten die alte Fassung. Das ist der leiseste Fehler in diesem
> Aufbau.

---

## Wenn etwas nicht geht

**Das Taskpane bleibt weiß.** Fast immer sendet der Webserver `X-Frame-Options`. Outlook
zeigt das Pane in einem Rahmen fremder Herkunft an; mit dieser Kopfzeile lädt es nicht,
und es erscheint keine Fehlermeldung. Prüfen:

```bash
curl -sI https://<ihre-ablage>/ | grep -i x-frame-options
```

Kommt eine Zeile zurück, muss sie für diesen Pfad entfernt werden — die vhost-Vorlage
zeigt, wie.

**„TANSS ist nicht erreichbar" im Pane.** Zwei Ursachen sind häufig: Die Adresse in
`apiBase` stimmt nicht (fehlt der Pfad `/backend`?), oder der Browser lehnt den fremden
Ursprung ab. Beides misst der Generator unter **Prüfen** nach.

**Die Schaltflächen erscheinen nicht.** Nach der Bereitstellung können 24 bis 72 Stunden
vergehen. Outlook neu starten. Erscheinen sie dann immer noch nicht, ist die Zuweisung
im Admin Center zu prüfen.

**Die Schaltflächen erscheinen, das Pane lädt nicht.** Der Arbeitsplatz erreicht die
Ablage nicht — typisch außerhalb des Firmennetzes ohne VPN, und auf Mobilgeräten.

**Der Termin findet seinen Einsatz nicht.** Die Seite sagt, worauf die Zuordnung beruht.
Steht dort ein Hinweis auf den Zeitabgleich, konnten die Kopplungsangaben nicht gelesen
werden; dann beruht die Zuordnung auf Beginn und Dauer und ist eine Faustregel. Bei
mehreren Treffern wird bewusst nicht geraten, sondern gefragt.

**Alles auf einmal prüfen:** die Schaltfläche **Prüfen** im Manifest-Generator. Sie ist
lesend, verlangt keine Anmeldung und ändert nichts.

---

## Entwicklung

```bash
npm test                                    # JavaScript, eingebauter Testläufer von Node
cd tools/ManifestGenerator && dotnet build  # der Generator
```

`package.json` bringt **keine Abhängigkeit und keinen Bauschritt** mit. Sie erklärt die
Dateien zu ES-Modulen, damit der Testläufer genau die Dateien laden kann, die auch der
Browser lädt. Dass sie diese Gestalt behält, wird mitgeprüft.

Die einzige mitgelieferte Fremdbibliothek ist `@azure/msal-browser`. Sie liegt mit
fester Fassung und SHA-256 in `taskpane/vendor/CHECKSUMS`; ein Test vergleicht bei jedem
Lauf.

---

## Lizenz

MIT. Siehe [LICENSE](LICENSE).

Zwei Dinge sind davon ausgenommen beziehungsweise gesondert lizenziert: das
Unternehmenszeichen in `taskpane/assets/logo.png` und die mitgelieferte Fremdbibliothek.
Beides steht in [NOTICE.md](NOTICE.md) — dort steht auch, wie das Zeichen für eine eigene
Installation ersetzt oder entfernt wird.
