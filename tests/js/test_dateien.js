/**
 * Die Waechter ueber die ausgelieferten Dateien.
 *
 * Durch dieses Taskpane laufen Betreff, vollstaendiger Klartext, komplettes MIME und ein
 * Anmeldetoken. Diese Daten duerfen keinen weiteren Empfaenger bekommen - nicht durch
 * eine nachtraeglich eingebundene Bibliothek, nicht durch einen Fehlersammler und nicht
 * durch eine Zeile, die in zwei Jahren jemand in guter Absicht hinzufuegt.
 *
 * Die Faelle hier lesen DATEIEN, keinen laufenden Code. Sie sind die Sperre, die auch
 * dann noch greift, wenn niemand mehr weiss, warum sie einmal eingezogen wurde. Jeder
 * einzelne deckt einen Weg ab, auf dem sich die Zusagen des Handbuchs lautlos
 * unterlaufen liessen.
 */

import test from "node:test";
import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFileSync, readdirSync, statSync } from "node:fs";
import { dirname, join, relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const WURZEL = resolve(dirname(fileURLToPath(import.meta.url)), "..", "..");
const TASKPANE = join(WURZEL, "taskpane");
const VENDOR = join(TASKPANE, "vendor");

/**
 * Die vier Herkuenfte, mit denen das Pane sprechen darf - und keine fuenfte.
 *
 * Die vierte ist nicht gewaehlt, sondern gemessen: office.js laedt fuer den Win32-Host
 * MicrosoftAjax von ajax.aspnetcdn.com nach. Fehlt der Eintrag, wird die Datei abgewiesen,
 * die Initialisierung kommt nicht zum Ende und `Office.onReady` meldet sich nie - ohne
 * jede Fehlermeldung im Pane. Genau daran hat die erste Auslieferung gehangen.
 */
const ERLAUBT = new Set([
  "https://appsforoffice.microsoft.com",
  "https://ajax.aspnetcdn.com",
  "https://login.microsoftonline.com",
  "https://graph.microsoft.com",
]);

/**
 * Namentlich gesperrt: verbreitete Fehlersammler und Skript-Auslieferungsnetze. Wer
 * eines davon einbindet, verschiebt die Vertraulichkeit dieses Panes zu einem Dritten -
 * und zwar so beilaeufig, dass es in keiner Durchsicht auffiele. Die Liste steht hier,
 * damit eine Wiedereinfuehrung nicht als "nur ein Host mehr" durchgeht.
 */
const VERBRANNT = ["sentry.io", "ingest.sentry", "cdn.jsdelivr.net",
  "cdnjs.cloudflare.com", "unpkg.com", "googleapis.com", "gstatic.com"];

/**
 * Fuer Beispiele reservierte Namensraeume. Sie gehoeren niemandem und loesen nirgends
 * auf - eine Adresse darunter kann kein Empfaenger sein. Sie muss aber vorkommen
 * duerfen: als Platzhalter in einem Eingabefeld, damit der Techniker die erwartete Form
 * sieht.
 */
const BEISPIEL_ENDUNGEN = [".example", ".example.de", ".example.com", ".example.net",
  ".example.org", ".invalid", ".test", ".localhost"];

/** Alles, was eine Zeichenkette zu Markup macht - jeder Eintrag ein eigener Weg. */
const MARKUP_WEGE = ["innerHTML", "outerHTML", "insertAdjacentHTML", "document.write",
  "createContextualFragment", "srcdoc", "eval(", "new Function("];

const URL_MUSTER = /https?:\/\/[A-Za-z0-9._%+-]+/g;
const JAVASCRIPT_URL = /["'(]\s*javascript:/i;

/** Die Beispielkonfiguration ist Anschauungsmaterial und wird gesondert geprueft. */
const BEISPIELKONFIGURATION = "config.example.json";

/* ------------------------------------------------------------------ Hilfsmittel */

function dateienUnter(ordner, endungen, ausnahme = () => false) {
  const gefunden = [];
  for (const name of readdirSync(ordner)) {
    const pfad = join(ordner, name);
    if (ausnahme(pfad)) continue;
    if (statSync(pfad).isDirectory()) {
      gefunden.push(...dateienUnter(pfad, endungen, ausnahme));
    } else if (endungen.some((endung) => name.endsWith(endung))) {
      gefunden.push(pfad);
    }
  }
  return gefunden;
}

/** Alle Dateien des Panes, die WIR geschrieben haben. */
function eigeneDateien(endungen = [".js", ".html", ".css", ".json"]) {
  // Auf Pfadsegmente pruefen, nicht auf Teilzeichenketten: Laege das Repository eines
  // Tages in einem Ordner, der "vendor" im Namen traegt, schloesse eine schlichte
  // Teilpruefung ALLE Dateien aus - und der Waechter bestuende stillschweigend ueber
  // nichts.
  const istVendor = (pfad) => relative(TASKPANE, pfad).split(/[\\/]/).includes("vendor");
  return dateienUnter(TASKPANE, endungen,
    (pfad) => istVendor(pfad) || pfad.endsWith(BEISPIELKONFIGURATION)).sort();
}

const lies = (pfad) => readFileSync(pfad, "utf8");
const kurz = (pfad) => relative(TASKPANE, pfad).replace(/\\/g, "/");

/**
 * Der Quelltext ohne Kommentare - nur er kann etwas abrufen.
 *
 * Die Zusage lautet "es wird kein fremder Host ANGESPROCHEN", nicht "das Wort kommt
 * nirgends vor": Ein Kommentar, der eine Adresse nennt, um sie zu erklaeren oder
 * auszuschliessen, ruft nichts auf.
 */
function ohneKommentare(pfad) {
  return lies(pfad)
    .replace(/<!--[\s\S]*?-->/g, " ")
    .replace(/\/\*[\s\S]*?\*\//g, " ")
    // Das `(?<!:)` haelt das `//` in `https://` aus der Zeilenkommentarregel heraus.
    .replace(/(?<!:)\/\/.*$/gm, " ");
}

/** Zeigt diese Adresse auf einen fuer Beispiele reservierten Namen? */
function istBeispieladresse(url) {
  const host = url.split("//")[1].split("/")[0].split(":")[0].toLowerCase();
  return BEISPIEL_ENDUNGEN.some((endung) => host.endsWith(endung));
}

/** Zerlegt eine Inhaltsrichtlinie in `{Direktive: [Quellen]}`. */
function cspDirektiven(text) {
  const direktiven = {};
  for (const teil of text.split(";")) {
    const stuecke = teil.trim().split(/\s+/).filter(Boolean);
    if (stuecke.length > 0) direktiven[stuecke[0]] = stuecke.slice(1);
  }
  return direktiven;
}

/** Der Inhalt des `<meta>`-Elements mit der Inhaltsrichtlinie. */
function metaRichtlinie() {
  const html = lies(join(TASKPANE, "index.html"));
  const beginn = html.indexOf('http-equiv="Content-Security-Policy"');
  assert.ok(beginn > 0, "index.html traegt keine Inhaltsrichtlinie");
  return html.slice(beginn).split('content="')[1].split('"')[0];
}

/* -------------------------------------------------------- Gibt es etwas zu pruefen */

test("es gibt ueberhaupt etwas zu pruefen", () => {
  // Ein Suchlauf ueber null Dateien ist gruen und sagt nichts. Verschiebt jemand
  // taskpane/ oder benennt es um, soll das hier auffallen und nicht stillschweigend
  // bestehen.
  const dateien = eigeneDateien();
  assert.ok(dateien.length >= 10, `nur ${dateien.length} Dateien gefunden - stimmt der Pfad?`);

  const namen = new Set(dateien.map((pfad) => pfad.split(/[\\/]/).pop()));
  for (const pflicht of ["index.html", "main.js", "ui.js", "api.js", "auth.js", "office.js"]) {
    assert.ok(namen.has(pflicht), `${pflicht} fehlt`);
  }
});

/* ------------------------------------------------------- Kein Markup aus Text */

for (const weg of MARKUP_WEGE) {
  test(`das Pane baut kein Markup ueber ${weg}`, () => {
    // Firmen-, Ticket- und Absendernamen kommen aus TANSS und damit aus einem System,
    // in das Kunden schreiben. EINE Zuweisung mit einem Firmennamen ist eine
    // Skriptausfuehrung im Pane - und das Pane haelt ein Anmeldetoken.
    const fundstellen = [];
    for (const pfad of eigeneDateien([".js", ".html", ".css"])) {
      lies(pfad).split(/\r?\n/).forEach((zeile, i) => {
        if (zeile.includes(weg)) fundstellen.push(`${kurz(pfad)}:${i + 1}: ${zeile.trim()}`);
      });
    }
    assert.deepEqual(fundstellen, [], `${weg} gefunden`);
  });
}

test("es gibt keine javascript:-Adresse", () => {
  const fundstellen = [];
  for (const pfad of eigeneDateien([".js", ".html", ".css"])) {
    lies(pfad).split(/\r?\n/).forEach((zeile, i) => {
      if (JAVASCRIPT_URL.test(zeile)) fundstellen.push(`${kurz(pfad)}:${i + 1}`);
    });
  }
  assert.deepEqual(fundstellen, []);
});

test("der eine Erzeuger setzt textContent", () => {
  const quelle = lies(join(TASKPANE, "js", "ui.js"));
  assert.ok(quelle.includes("textContent"));
  assert.ok(quelle.includes("export function el("));
});

test("die Seiten bauen ihre Knoten nicht selbst", () => {
  // Wer `document.createElement` selbst aufruft, umgeht den einen Erzeuger - und mit
  // ihm die Regel, die er traegt.
  const fundstellen = dateienUnter(join(TASKPANE, "js", "pages"), [".js"])
    .filter((pfad) => lies(pfad).includes("document.createElement"))
    .map(kurz);
  assert.deepEqual(fundstellen, []);
});

/* ------------------------------------------------------------ Keine fremden Hosts */

test("im Pane steht keine fremde Adresse", () => {
  const fundstellen = [];
  for (const pfad of eigeneDateien()) {
    ohneKommentare(pfad).split(/\r?\n/).forEach((zeile, i) => {
      for (const treffer of zeile.match(URL_MUSTER) || []) {
        if (ERLAUBT.has(treffer.replace(/\/+$/, "")) || istBeispieladresse(treffer)) continue;
        fundstellen.push(`${kurz(pfad)}:${i + 1}: ${treffer}`);
      }
    });
  }
  assert.deepEqual(fundstellen, [], "fremde Adresse im Pane");
});

for (const host of VERBRANNT) {
  test(`der gesperrte Host ${host} taucht nicht auf`, () => {
    for (const pfad of eigeneDateien()) {
      assert.ok(!ohneKommentare(pfad).includes(host), `${host} steht in ${kurz(pfad)}`);
    }
  });
}

test("es gibt keinen Fehlersammler im Pane", () => {
  for (const pfad of eigeneDateien()) {
    const quelle = ohneKommentare(pfad);
    assert.ok(!quelle.includes("Sentry"), kurz(pfad));
    assert.ok(!quelle.includes("tracesSampleRate"), kurz(pfad));
  }
});

test("die Aufrufe stehen nicht als feste Adresse im Quelltext", () => {
  // Das Add-in laeuft unter jedem Hostnamen, ohne dass etwas umgeschrieben werden muss -
  // und es gibt keine Stelle, an der jemand eine feste Basis unterschieben koennte.
  const dateien = [join(TASKPANE, "js", "api.js"),
    ...dateienUnter(join(TASKPANE, "js", "pages"), [".js"])];
  for (const pfad of dateien) {
    for (const treffer of ohneKommentare(pfad).match(URL_MUSTER) || []) {
      assert.ok(istBeispieladresse(treffer),
        `${kurz(pfad)} enthaelt die feste Adresse ${treffer}. Aufrufe gehoeren ueber die `
        + "Konfiguration, nicht in den Quelltext.");
    }
  }
});

test("die Beispielkonfiguration nennt nur Beispieladressen", () => {
  // Die Datei wird mit dem Repository veroeffentlicht. Eine dort vergessene Adresse
  // verriete, wer dieses Werkzeug einsetzt, und liesse sich nicht zurueckholen -
  // Beispieldateien werden kopiert, nicht gelesen.
  const pfad = join(TASKPANE, BEISPIELKONFIGURATION);
  const fundstellen = [];
  lies(pfad).split(/\r?\n/).forEach((zeile, i) => {
    for (const treffer of zeile.match(URL_MUSTER) || []) {
      if (!istBeispieladresse(treffer)) fundstellen.push(`${BEISPIELKONFIGURATION}:${i + 1}: ${treffer}`);
    }
  });
  assert.deepEqual(fundstellen, []);
});

/* ------------------------------------------------------------------- Die Seite */

test("die Seite laedt ein Stylesheet und keine Schrift", () => {
  // Jede Schriftdatei ist ein weiterer Abruf vor der ersten Anzeige und, kaeme sie von
  // einem fremden Host, ein weiterer Eintrag in der Inhaltsrichtlinie.
  const html = lies(join(TASKPANE, "index.html"));
  assert.equal((html.match(/<link rel="stylesheet"/g) || []).length, 1);
  assert.ok(!lies(join(TASKPANE, "app.css")).includes("@font-face"));
});

test("die Seite traegt kein BOM", () => {
  // Ein BOM macht aus der ersten Zeile Text VOR dem Dokumenttyp; manche Hosts schalten
  // dann in einen Kompatibilitaetsmodus.
  const roh = readFileSync(join(TASKPANE, "index.html"));
  assert.notEqual(roh[0], 0xef);
  assert.ok(roh.toString("utf8").trimStart().startsWith("<!doctype html>"));
});

/* -------------------------------------------------------------- Die Richtlinie */

test("die Richtlinie im meta-Element nennt genau die noetigen Herkuenfte", () => {
  const direktiven = cspDirektiven(metaRichtlinie());
  assert.deepEqual(direktiven["default-src"], ["'self'"]);
  // ajax.aspnetcdn.com steht hier, weil office.js von dort MicrosoftAjax nachlaedt. Ohne
  // den Eintrag bleibt das Pane stumm stehen - der Fall ist gemessen, nicht angenommen.
  assert.deepEqual(direktiven["script-src"],
    ["'self'", "https://appsforoffice.microsoft.com", "https://ajax.aspnetcdn.com"]);
  // `https:` steht fuer die TANSS-Instanz, deren Adresse erst in der config.json der
  // jeweiligen Installation steht. Verschaerft wird die Direktive vom Ablageort.
  assert.deepEqual(direktiven["connect-src"],
    ["https:", "https://login.microsoftonline.com", "https://graph.microsoft.com"]);
  assert.deepEqual(direktiven["object-src"], ["'none'"]);
  assert.deepEqual(direktiven["form-action"], ["'none'"]);
});

test("die Richtlinie im meta-Element setzt keine frame-ancestors", () => {
  // In einem meta-Element wird die Anweisung vom Browser IGNORIERT. Kaeme sie nur von
  // dort, waere das Pane in jede fremde Seite einbettbar - und der Entwurf saehe
  // trotzdem richtig aus.
  assert.ok(!metaRichtlinie().includes("frame-ancestors"));
});

test("die Ablagevorlage verschaerft die Verbindungsdirektive", () => {
  // Was das meta-Element offen lassen muss, schliesst der Ablageort. Die Vorlage nennt
  // dafuer einen Platzhalter statt einer Adresse: Wer ihn stehen laesst, bekommt eine
  // Richtlinie, die JEDEN Aufruf an TANSS blockiert - und zwar leise, denn der Browser
  // meldet es nur in seiner eigenen Konsole.
  const vorlage = join(WURZEL, "deploy", "apache-addon-vhost.conf.example");
  const text = lies(vorlage);
  const zeile = text.slice(text.indexOf("Content-Security-Policy")).split("\n")[0];
  const direktiven = cspDirektiven(zeile.split('"')[1]);

  assert.ok(direktiven["connect-src"].includes("<tanss-url>"),
    "die Vorlage nennt die TANSS-Adresse nicht in connect-src");
  assert.ok(!direktiven["connect-src"].includes("https:"),
    "die Vorlage laesst connect-src offen - dort ist die Adresse aber bekannt");
  assert.ok("frame-ancestors" in direktiven,
    "ohne frame-ancestors in der Kopfzeile ist das Pane in jede fremde Seite einbettbar");
  for (const herkunft of ["https://*.office.com", "https://*.outlook.com"]) {
    assert.ok(direktiven["frame-ancestors"].includes(herkunft), `${herkunft} fehlt`);
  }
});

/* ------------------------------------------------------------ Die Fremdbibliothek */

test("jede mitgelieferte Fremddatei stimmt mit ihrer Pruefsumme", () => {
  // Sie laeuft mit vollen Rechten im Pane und sieht jedes Token, das dort durchgeht.
  // Ein untergeschobener Austausch - versehentlich durch eine
  // Zeilenende-Umwandlung oder absichtlich - waere der wirksamste Angriff auf dieses
  // Werkzeug und von aussen nicht zu bemerken.
  const erwartet = new Map();
  for (const zeile of lies(join(VENDOR, "CHECKSUMS")).split(/\r?\n/)) {
    const text = zeile.trim();
    if (!text || text.startsWith("#")) continue;
    const treffer = /^([0-9a-f]{64})\s+(\S+)$/.exec(text);
    assert.ok(treffer, `unlesbare Zeile in CHECKSUMS: ${text}`);
    erwartet.set(treffer[2], treffer[1]);
  }
  assert.ok(erwartet.size > 0, "CHECKSUMS enthaelt keine einzige Pruefsumme");

  for (const [name, summe] of erwartet) {
    const gemessen = createHash("sha256").update(readFileSync(join(VENDOR, name))).digest("hex");
    assert.equal(gemessen, summe,
      `${name} weicht von der hinterlegten Pruefsumme ab. Entweder wurde die Datei `
      + "ausgetauscht, oder sie wurde beim Auschecken umkodiert.");
  }

  // Eine liegengebliebene Altfassung wuerde vom eigenen Ursprung ausgeliefert und waere
  // von aussen erreichbar, obwohl sie niemand mehr prueft.
  const vorhanden = readdirSync(VENDOR).filter((name) => name !== "CHECKSUMS");
  const ungedeckt = vorhanden.filter((name) => !erwartet.has(name));
  assert.deepEqual(ungedeckt, [], "diese Dateien liegen in vendor/, stehen aber in keiner Zeile");
});

test("genau eine Fassung der Fremdbibliothek liegt bei, und die Seite laedt sie", () => {
  const fassungen = readdirSync(VENDOR).filter((name) => name.startsWith("msal-browser-"));
  assert.equal(fassungen.length, 1, `mehr als eine Fassung: ${fassungen}`);
  // Kein "@latest", kein "^4": Eine bewegliche Fassung waere dieselbe Luecke wie ein
  // Auslieferungsnetz, nur langsamer.
  assert.match(fassungen[0], /^msal-browser-\d+\.\d+\.\d+\.min\.js$/);
  // Der Dateiname ist in index.html verdrahtet. Zeigt er auf eine andere Datei, ist die
  // Pruefsumme richtig und trotzdem wertlos.
  assert.ok(lies(join(TASKPANE, "index.html")).includes(`src="vendor/${fassungen[0]}"`));
  assert.deepEqual(readdirSync(VENDOR).filter((name) => name.endsWith(".map")), [],
    "eine Quellkarte gehoert nicht auf den Server");
});

test("die Lizenz der Fremdbibliothek liegt daneben", () => {
  const wortlaut = lies(join(VENDOR, "LICENSE-msal"));
  assert.ok(wortlaut.includes("MIT"));
  assert.ok(wortlaut.includes("Microsoft"));
});

/* ------------------------------------------------------------ Keine Abhaengigkeit */

test("es gibt keine Abhaengigkeit und keinen Bauschritt", () => {
  // Jede Abhaengigkeit waere ein weiterer Herausgeber von Code, der im Pane laeuft - wo
  // Betreff, Klartext und ein Token durchgehen. Die package.json ist zulaessig, aber
  // nur in einer Gestalt: ohne Abhaengigkeit und ohne Bauschritt.
  for (const verboten of ["package-lock.json", "yarn.lock", "pnpm-lock.yaml", "node_modules"]) {
    assert.ok(!existiert(join(WURZEL, verboten)), `${verboten} liegt im Repository`);
  }

  const paket = JSON.parse(lies(join(WURZEL, "package.json")));
  for (const feld of ["dependencies", "devDependencies", "peerDependencies",
    "optionalDependencies"]) {
    assert.ok(!paket[feld] || Object.keys(paket[feld]).length === 0,
      `package.json fuehrt ${feld} - damit gaebe es etwas zu installieren`);
  }
  assert.deepEqual(Object.keys(paket.scripts || {}), ["test"],
    "ein Bauschritt gehoert nicht dazu - ausgeliefert wird, was hier liegt");
  assert.equal(paket.type, "module",
    'ohne "type": "module" laedt der Testlaeufer die Dateien als CommonJS');
});

function existiert(pfad) {
  try {
    statSync(pfad);
    return true;
  } catch {
    return false;
  }
}

/* --------------------------------------------------------------- Das Manifest */

test("jedes Symbol, auf das die Manifestvorlage zeigt, liegt auch da", () => {
  // Ein Manifest, das auf eine fehlende Datei zeigt, ist gueltig und sieht richtig aus.
  // Outlook zeigt dann einen Platzhalter im Menueband - und niemand kommt darauf, dass
  // die Ursache eine Datei ist, die beim Umbenennen vergessen wurde.
  const vorlage = lies(join(WURZEL, "tools", "ManifestGenerator", "manifest.xml.tmpl"));
  const verweise = [...vorlage.matchAll(/\$\{base_url\}\/(assets\/[A-Za-z0-9._-]+)/g)]
    .map((treffer) => treffer[1]);

  assert.ok(verweise.length >= 5, `nur ${verweise.length} Symbolverweise - stimmt die Vorlage?`);
  for (const verweis of new Set(verweise)) {
    assert.ok(existiert(join(TASKPANE, verweis)),
      `Die Manifestvorlage zeigt auf ${verweis}, dort liegt aber nichts.`);
  }
});

test("das Menue der Mailansicht traegt zwei unterscheidbare Eintraege", () => {
  // Zwei Eintraege mit demselben Symbol oder derselben Adresse waeren eine Auswahl, die
  // keine ist - und das faellt erst auf, wenn jemand beide ausprobiert.
  const vorlage = lies(join(WURZEL, "tools", "ManifestGenerator", "manifest.xml.tmpl"));

  const menues = (vorlage.match(/xsi:type="Menu"/g) || []).length;
  assert.equal(menues, 1, "erwartet wird genau ein Aufklappmenue");

  const eintraege = [...vorlage.matchAll(/<Item id="([^"]+)"/g)].map((t) => t[1]);
  assert.equal(eintraege.length, 2, `Eintraege: ${eintraege}`);

  const inMenu = vorlage.slice(vorlage.indexOf("<Items>"), vorlage.indexOf("</Items>"));
  const symbole = [...inMenu.matchAll(/resid="(i[A-Za-z]+\d+)"/g)].map((t) => t[1]);
  assert.equal(new Set(symbole).size, symbole.length,
    `zwei Eintraege teilen sich ein Symbol: ${symbole}`);

  const ziele = [...inMenu.matchAll(/<SourceLocation resid="([^"]+)"/g)].map((t) => t[1]);
  assert.equal(new Set(ziele).size, ziele.length, `zwei Eintraege zeigen auf ${ziele}`);
});

test("die Beschriftungen im Menueband tragen echte Umlaute", () => {
  // Sie liest der Benutzer. Die Umlautregel dieses Repos gilt fuer Quelltextkommentare,
  // nicht fuer Anzeigetexte - dort sieht "anhaengen" nach einem Fehler aus.
  const vorlage = lies(join(WURZEL, "tools", "ManifestGenerator", "manifest.xml.tmpl"));
  const texte = [...vorlage.matchAll(/DefaultValue="([^"$]+)"/g)]
    .map((treffer) => treffer[1])
    .filter((text) => !text.startsWith("http") && !text.includes("assets/"));

  assert.ok(texte.length >= 5, "kaum Anzeigetexte gefunden - stimmt die Vorlage?");
  for (const text of texte) {
    for (const ersatz of ["ae", "oe", "ue"]) {
      assert.ok(!new RegExp(`[a-z]${ersatz}[a-z]`).test(text),
        `Anzeigetext mit Umlautersatz: "${text}"`);
    }
  }
});

test("das Startprotokoll wird vor allem anderen geladen", () => {
  // Sein Zuhoerer meldet Verstoesse gegen die Inhaltsrichtlinie erst ab dem Augenblick, in
  // dem er haengt. Rutschte diese Zeile in die alphabetische Ordnung darunter, fehlte auf
  // der Diagnoseseite genau das, wofuer es sie gibt - und niemand saehe, dass etwas fehlt.
  const quelle = lies(join(TASKPANE, "js", "main.js"));
  const importe = [...quelle.matchAll(/^import .*? from "([^"]+)";$/gm)].map((t) => t[1]);
  assert.equal(importe[0], "./boot.js", `erste Import-Zeile ist ${importe[0]}`);
});

test("keine Seite ruft eine Route auf, die der Verteiler nicht kennt", () => {
  // Frueher lag hinter dem Verteiler ein eigener Dienst mit eigenen Routen. Der ist weg,
  // seine Pfade sind es nicht ueberall gewesen: Die Diagnoseseite fragte weiter nach
  // `healthz` und `readyz` und zeigte dafuer zwei erfundene Fehler an - ausgerechnet die
  // Seite, die man oeffnet, WEIL etwas klemmt. Ein toter Pfad faellt sonst erst dem
  // Benutzer auf, und ihm sagt er nichts.
  const verteiler = lies(join(TASKPANE, "js", "api.js"));
  const dateien = [join(TASKPANE, "js", "main.js")];
  for (const name of readdirSync(join(TASKPANE, "js", "pages"))) {
    if (name.endsWith(".js")) dateien.push(join(TASKPANE, "js", "pages", name));
  }

  let geprueft = 0;
  for (const datei of dateien) {
    const quelle = lies(datei);
    for (const treffer of quelle.matchAll(/api\.(?:get|post|put)\(\s*"([^"]+)"/g)) {
      const pfad = treffer[1];
      geprueft += 1;
      const alsMuster = `^${pfad.replace(/\//g, "\\/")}$`;
      assert.ok(verteiler.includes(alsMuster) || verteiler.includes(pfad),
        `${relative(WURZEL, datei)} ruft "${pfad}" auf, der Verteiler kennt den Pfad nicht.`);
    }
  }

  // Ein Waechter ohne Gegenstand sieht auch dann nichts, wenn er kaputt ist.
  assert.ok(geprueft > 0, "kein einziger Routenaufruf gefunden - der Fall prueft nichts");
});

test("der fruehe Zuhoerer steht vor office.js und ist kein Modul", () => {
  // Beides ist Bedingung, nicht Geschmack. Ein Modul liefe zurueckgestellt - also erst
  // nach office.js -, und genau die Meldungen, um die es geht, fallen waehrend dessen
  // Start. Steht die Zeile hinter office.js oder traegt sie type="module", misst das Pane
  // eine Leere und meldet sie als Entwarnung.
  const seite = lies(join(TASKPANE, "index.html"));

  const zuhoerer = seite.indexOf('src="js/boot-early.js"');
  const office = seite.indexOf("appsforoffice.microsoft.com/lib/1/hosted/office.js");
  assert.ok(zuhoerer > -1, "js/boot-early.js wird in der index.html nicht geladen");
  assert.ok(office > -1, "office.js wird in der index.html nicht geladen");
  assert.ok(zuhoerer < office, "der fruehe Zuhoerer steht hinter office.js");

  const zeile = seite.slice(seite.lastIndexOf("<script", zuhoerer), zuhoerer);
  assert.ok(!zeile.includes("type=\"module\""), "der fruehe Zuhoerer ist als Modul eingebunden");
});

test("der fruehe Zuhoerer bleibt klein und ohne Abhaengigkeit", () => {
  // Was hier scheitert, scheitert vor jeder Fehlerbehandlung: Zu diesem Zeitpunkt gibt es
  // weder eine Oberflaeche noch ein Auffangnetz. Die Datei darf deshalb nicht wachsen und
  // nichts nachladen.
  const quelle = lies(join(TASKPANE, "js", "boot-early.js"));
  assert.ok(quelle.length < 2500, `boot-early.js ist ${quelle.length} Zeichen gross`);
  assert.ok(!/\bimport\b|\brequire\(|\bfetch\(/.test(quelle),
    "boot-early.js laedt etwas nach - das darf es nicht");
});

test("meta-Element und Ablagevorlage nennen dieselben Skript-Herkuenfte", () => {
  // Die Liste steht an zwei Orten, und beide muessen gelten: Das meta-Element traegt sie
  // auf jedem Hoster, die Kopfzeile zusaetzlich dort, wo man Kopfzeilen setzen kann. Je
  // Direktive gewinnt die STRENGERE - eine Herkunft, die nur in einer der beiden steht,
  // ist damit wirkungslos, und zwar leise. Genau so ist ajax.aspnetcdn.com beim ersten
  // Mal ueberhaupt erst aufgefallen: als stummer Stillstand ohne Fehlermeldung.
  const ausMeta = cspDirektiven(metaRichtlinie())["script-src"];

  const vorlage = lies(join(WURZEL, "deploy", "apache-addon-vhost.conf.example"));
  const zeile = vorlage.slice(vorlage.indexOf("Content-Security-Policy")).split("\n")[0];
  const ausKopfzeile = cspDirektiven(zeile.split('"')[1])["script-src"];

  assert.deepEqual(ausKopfzeile, ausMeta,
    "die beiden Fassungen der Skript-Direktive sind auseinandergelaufen");
});

test("jede Skript-Herkunft der Richtlinie wird auch wirklich gebraucht", () => {
  // Eine Herkunft, die niemand mehr laedt, ist eine offene Tuer ohne Zweck. Sie muss
  // entweder in der index.html vorkommen oder im Quelltext begruendet sein - sonst ist
  // sie ein Rest, den jemand einzutragen vergessen hat auszutragen.
  const seite = lies(join(TASKPANE, "index.html"));
  const herkuenfte = cspDirektiven(metaRichtlinie())["script-src"]
    .filter((eintrag) => eintrag.startsWith("https://"));

  for (const herkunft of herkuenfte) {
    const host = herkunft.replace("https://", "");
    assert.ok(seite.includes(host),
      `${herkunft} steht in der Richtlinie, kommt in der index.html aber nicht vor`);
  }
});

test("der Firmenwechsel raeumt alles ab, was zur alten Firma gehoerte", () => {
  // Die Maske hat keinen Unit-Test - es gibt kein DOM. Die dauerhafte Sperre steht
  // deshalb im Repository und ist dort geprueft. Dieser Waechter haelt nur die zweite
  // Haelfte fest: dass beim Firmenwechsel auch das SICHTBARE verschwindet. Genau daran
  // hing der gemeldete Fehler - `state.remitter` wurde verworfen, die Trefferliste der
  // alten Firma blieb stehen und anklickbar.
  const quelle = lies(join(TASKPANE, "js", "pages", "ticket-neu.js"));
  const anfang = quelle.indexOf("function pickCompany(");
  assert.ok(anfang > -1, "pickCompany gibt es nicht mehr");
  const koerper = quelle.slice(anfang, quelle.indexOf("\n  }", anfang));

  for (const pflicht of ["state.remitter = null", "state.similar = []", "clearRemitterSearch()"]) {
    assert.ok(koerper.includes(pflicht), `pickCompany raeumt nicht ab: ${pflicht} fehlt`);
  }
});

test("die Meldersuche prueft die Firma nach dem Warten erneut", () => {
  // Eine noch laufende Suche liefert nach dem Firmenwechsel nach. Ihr Ergebnis gehoert
  // zur ALTEN Firma und darf in der neuen Maske weder erscheinen noch anklickbar sein.
  // Das `signal` faengt das nicht: Es faellt beim Seitenwechsel, nicht beim
  // Firmenwechsel.
  const quelle = lies(join(TASKPANE, "js", "pages", "ticket-neu.js"));
  const anfang = quelle.indexOf("async function searchEmployees(");
  assert.ok(anfang > -1, "searchEmployees gibt es nicht mehr");
  const koerper = quelle.slice(anfang, quelle.indexOf("\n  }", anfang));

  assert.ok(/const\s+\w+\s*=\s*state\.company\.id/.test(koerper),
    "die Firma wird nicht festgehalten, sondern beim Auswerten neu nachgeschlagen");
  assert.ok(koerper.includes("await"), "ohne Warten braeuchte es die Pruefung nicht");

  const nachDemWarten = koerper.slice(koerper.indexOf("await"));
  assert.ok(/state\.company\.id\s*!==/.test(nachDemWarten),
    "nach dem Warten wird die Firma nicht gegen die festgehaltene geprueft");
});

test("das Pane ruft nur Flaechen auf, die seine Tokenart erreicht", () => {
  // TANSS bindet jede Schnittstellenflaeche an eine Rolle, und ein Anmeldetoken traegt
  // genau eine. Eine gewoehnliche Anmeldung ergibt die Rolle des Fachzugriffs - damit
  // sind /api/v1/** (einschliesslich /api/v1/admin/**, wo "admin" ein Namensteil ist und
  // keine Rollenforderung) erreichbar, die Integrationsflaeche /api/tanss.x/v1/** und die
  // ERP-Flaeche /api/erp/v1/** dagegen NICHT.
  //
  // Das ist kein gelegentlicher Ausfall, sondern eine Eigenschaft: Ein Aufruf dorthin
  // wird mit dem Token eines Technikers IMMER abgewiesen. Genau daran hing die Meldung
  // "Die Technikerliste war nicht abrufbar" - zweimal, ueber Tage.
  //
  // Der Waechter sieht nur die Aufrufstellen an. Dass client.js den Praefix als KONSTANTE
  // kennt, ist richtig und bleibt: Die Regel, dort niemals loggedInUserId anzuhaengen,
  // gilt weiter fuer den Fall, dass jemand die Flaeche einmal mit passendem Token nutzt.
  const unerreichbar = ["/api/tanss.x/", "/api/erp/"];
  const fundstellen = [];

  for (const datei of ["repository.js", "session.js"]) {
    const pfad = join(TASKPANE, "js", "tanss", datei);
    ohneKommentare(pfad).split(/\r?\n/).forEach((zeile, i) => {
      for (const flaeche of unerreichbar) {
        if (zeile.includes(`"${flaeche}`)) fundstellen.push(`${datei}:${i + 1}: ${flaeche}`);
      }
    });
  }

  assert.deepEqual(fundstellen, [],
    "Aufruf einer Flaeche, die eine andere Tokenart verlangt - er wird immer abgewiesen");
});

test("die Ticketliste wird nie per GET gerufen", () => {
  // Unter /api/v1/tickets sind nur POST und PUT definiert - ein GET auf diese Route gibt
  // es nicht. Der frueher dort vermutete meta-Block mit der Feldsteuerung kostete bei
  // jedem Oeffnen der Maske eine Umlaufzeit und landete zuverlaessig im Rueckfall. Die
  // Maske meldete deshalb IMMER, die Listen staemmten aus dem Rueckfall - und schob damit
  // einer Gegenstelle zu, was hier falsch gebaut war.
  const quelle = ohneKommentare(join(TASKPANE, "js", "tanss", "repository.js"));
  assert.ok(!/client\.get\(\s*["'`]\/api\/v1\/tickets\/["'`]/.test(quelle),
    "GET auf /api/v1/tickets/ - diese Route gibt es nicht");
});

test("die Prioritaetsstufen tragen keine Worte", () => {
  // Ob 1 oder 9 "hoch" bedeutet, geht aus der TANSS-Schnittstelle nicht hervor: Verglichen
  // wird nirgends groesser/kleiner, und die einzige Richtungsangabe der Dokumentation
  // gehoert dem Rueckruf, nicht dem Ticket. Eine Beschriftung wie "hoch" oder "niedrig"
  // waere deshalb eine Behauptung - und eine verkehrt herum eingebaute Skala setzte jedes
  // dringende Ticket auf die falsche Stufe, ohne dass es auffiele.
  const quelle = lies(join(TASKPANE, "js", "pages", "ticket-neu.js"));
  const block = quelle.slice(quelle.indexOf("const PRIORITIES"), quelle.indexOf("});", quelle.indexOf("const PRIORITIES")));

  for (const wort of ["hoch", "niedrig", "dringend", "kritisch", "normal", "sofort"]) {
    assert.ok(!new RegExp(wort, "i").test(block),
      `Die Prioritaetsliste behauptet eine Richtung: "${wort}"`);
  }
});
