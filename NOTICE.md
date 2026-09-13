# Hinweise zu Marken und mitgelieferten Dateien

Der Quelltext dieses Repositories steht unter der MIT-Lizenz (siehe `LICENSE`). Zwei
Dinge darin sind davon ausgenommen beziehungsweise gesondert lizenziert.

## Das Zeichen von ProNet Systems

`taskpane/assets/logo.png` ist das Unternehmenszeichen der ProNet Systems GmbH. Es ist
**nicht** Gegenstand der MIT-Lizenz.

Die MIT-Lizenz erlaubt jedem, den lizenzierten Inhalt zu verwenden, zu verändern und
weiterzuverbreiten. Für Quelltext ist genau das gewollt. Für ein Unternehmenszeichen
wäre es das Gegenteil: Es steht dafür, wer hinter einer Sache steht, und darf deshalb
nicht in fremdem Namen verwendet werden.

Wer dieses Add-in einsetzt, anpasst oder weitergibt, darf das Zeichen daher **nicht** so
verwenden, dass der Eindruck entsteht, ProNet Systems stehe hinter der eigenen Fassung.
Für eine eigene Installation ist die Datei schlicht zu ersetzen:

```
taskpane/assets/logo.png     288 x 100 Pixel, PNG mit Transparenz
```

Das Bild wird 25 Pixel hoch angezeigt; die Breite ergibt sich aus dem
Seitenverhältnis. Wird es ersetzt, gehört der Alternativtext in
`taskpane/i18n/de.js` unter `app.vendor` mit angepasst — er ist das, was ein
Screenreader vorliest.

Wer das Zeichen entfernen statt ersetzen möchte, löscht die Datei und nimmt in
`taskpane/js/main.js` in `buildShell()` die Zeile heraus, die es einhängt. Das Pane
funktioniert ohne.

## @azure/msal-browser

`taskpane/vendor/msal-browser-*.min.js` ist eine unveränderte Kopie von
`@azure/msal-browser` der Microsoft Corporation, ebenfalls unter der MIT-Lizenz. Der
Wortlaut liegt daneben in `taskpane/vendor/LICENSE-msal`.

Die Datei ist mit einer SHA-256-Prüfsumme in `taskpane/vendor/CHECKSUMS` festgehalten,
und ein Test vergleicht bei jedem Lauf. Sie läuft mit vollen Rechten im Taskpane und
sieht jedes Microsoft-Token, das dort durchgeht — ein untergeschobener Austausch wäre
der wirksamste Angriff auf dieses Werkzeug und von außen nicht zu bemerken.

## office.js

`office.js` wird vom Microsoft-CDN geladen und ist **nicht** Teil dieses Repositories.
Der Host erwartet die dort ausgelieferte Fassung; eine mitgelieferte Kopie würde bei
jedem Office-Update veralten.
