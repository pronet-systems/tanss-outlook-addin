/**
 * Beschaffung der geoeffneten E-Mail als RFC-822-Nachricht.
 *
 * Zwei Wege, zur Laufzeit entschieden, ein Ergebnis:
 *
 * (a) Postfach in Exchange Online - das TASKPANE holt die Nachricht selbst bei Microsoft
 *     Graph (`GET /me/messages/{id}/$value`) mit einem NAA-Access-Token auf delegiertes
 *     `Mail.Read`. Der Dienst braucht dadurch KEINERLEI Graph-Berechtigung, kein
 *     Microsoft-Geheimnis und hat keine Reichweite auf fremde Postfaecher. Das ist keine
 *     Betriebszusage, sondern eine Struktureigenschaft: das Graph-Token existiert nur im
 *     Browser des angemeldeten Technikers.
 *
 * (b) Postfach auf lokalem Exchange (`accountType === "enterprise"`) - Graph steht nicht zur
 *     Verfuegung. Die Nachricht wird aus Office.js-Bordmitteln zusammengesetzt, MIT den
 *     Anhaengen ueber `getAttachmentContentAsync` (Mailbox 1.8).
 *
 * Was Weg (b) NICHT kann, und was deshalb in der README steht: Es entsteht eine
 * Rekonstruktion, keine Kopie. `Received`-Kette, `DKIM-Signature` und `References` fehlen;
 * die `Message-ID` uebernehmen wir aus `item.internetMessageId`. Fuer Beweiszwecke ist die
 * so abgelegte Mail schwaecher als eine, die der Mailroboter importiert hat.
 *
 * Ergebnis beider Wege ist derselbe Umschlag wie in `api.js`, damit die Seiten nur eine
 * Fehlerbehandlung brauchen.
 */

import { T, formatBytes, t } from "../i18n/de.js";
import { getGraphToken } from "./auth.js";
import {
  convertToRestId,
  getAttachmentContent,
  getAttachments,
  getBody,
  getMessageContext,
  isEnterpriseMailbox,
  isSetSupported,
} from "./office.js";

const GRAPH_MESSAGE_URL = "https://graph.microsoft.com/v1.0/me/messages/";

/** Delegierte Berechtigung fuer den Graph-Weg. Mehr wird nie angefordert. */
const GRAPH_MAIL_SCOPE = "Mail.Read";

/** Aufschlag, mit dem Base64 und MIME-Rahmen die Rohgroesse aufblaehen (4/3 plus Koepfe). */
const BASE64_OVERHEAD = 1.4;

const CRLF = "\r\n";

/* ------------------------------------------------------------------- Umschlag */

function ok(data) {
  return { ok: true, data };
}

function fail(code, message = "") {
  return { ok: false, error: { code, message: message || T.errors[code] || T.errors.CLIENT_UNEXPECTED } };
}

/* ------------------------------------------------------------------ Hauptweg */

/**
 * Holt die geoeffnete Nachricht als `message/rfc822`.
 *
 * @param {{maxBytes?: number}} [options] Obergrenze aus `GET /api/me` (`limits.maxEmlBytes`).
 *        Wird sie ueberschritten, bricht dieser Aufruf ab, BEVOR etwas uebertragen wird -
 *        ein HTTP 413 nach 25 MiB Upload ist die unfreundlichste Art, es zu erfahren.
 * @returns {Promise<{ok: true, data: {blob: Blob, bytes: number, filename: string,
 *   subject: string, internetMessageId: string, source: "graph"|"officejs",
 *   attachmentCount: number, warnings: string[]}} | {ok: false, error: object}>}
 */
export async function fetchMessageMime(options = {}) {
  const maxBytes = Number(options.maxBytes) > 0 ? Number(options.maxBytes) : 0;

  let context;
  try {
    context = getMessageContext();
  } catch (error) {
    return fail(error && error.code ? error.code : "CLIENT_ITEM_MISSING");
  }
  if (!context.itemId) return fail("CLIENT_ITEM_MISSING");

  return isEnterpriseMailbox()
    ? buildFromOfficeJs(context, maxBytes)
    : fetchFromGraph(context, maxBytes);
}

/** Nur zur Anzeige: welcher Weg genommen wuerde. `#/diagnose` zeigt das. */
export function currentSource() {
  return isEnterpriseMailbox() ? "officejs" : "graph";
}

/* -------------------------------------------------------------- Weg (a) Graph */

async function fetchFromGraph(context, maxBytes) {
  let token;
  try {
    token = await getGraphToken(GRAPH_MAIL_SCOPE);
  } catch (error) {
    return fail(error && error.code ? error.code : "CLIENT_AUTH_FAILED");
  }

  let restId;
  try {
    restId = convertToRestId(context.itemId);
  } catch {
    return fail("CLIENT_OFFICE_FAILED");
  }

  let response;
  try {
    response = await fetch(`${GRAPH_MESSAGE_URL}${encodeURIComponent(restId)}/$value`, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${token}`,
        // Unveraenderliche Kennungen: eine verschobene Nachricht behaelt ihre Id.
        Prefer: 'IdType="ImmutableId"',
      },
      credentials: "omit",
      cache: "no-store",
    });
  } catch {
    return fail("CLIENT_NETWORK");
  }

  if (!response.ok) {
    // 401/403 heisst hier: die delegierte Berechtigung Mail.Read fehlt oder wurde
    // zurueckgezogen. Ein zweiter Versuch mit demselben Token brachte nichts.
    return fail(
      response.status === 401 || response.status === 403
        ? "CLIENT_AUTH_CONSENT_REQUIRED"
        : "CLIENT_GRAPH_FAILED",
    );
  }

  // Groesse VOR dem Einlesen pruefen, solange sie noch eine Kopfzeile und kein Speicher ist.
  const announced = Number(response.headers.get("Content-Length") || 0);
  if (maxBytes > 0 && announced > maxBytes) {
    return tooLarge(announced, maxBytes);
  }

  let blob;
  try {
    blob = await response.blob();
  } catch {
    return fail("CLIENT_GRAPH_FAILED");
  }
  if (maxBytes > 0 && blob.size > maxBytes) return tooLarge(blob.size, maxBytes);

  return ok({
    blob: new Blob([blob], { type: "message/rfc822" }),
    bytes: blob.size,
    filename: emlFilename(context.subject),
    subject: context.subject,
    internetMessageId: context.internetMessageId,
    source: "graph",
    attachmentCount: countRealAttachments(safeAttachments()),
    warnings: [],
  });
}

/* --------------------------------------------------- Weg (b) Bordmittel (RFC 822) */

async function buildFromOfficeJs(context, maxBytes) {
  const warnings = [T.mail.reconstructed];
  const attachments = safeAttachments();

  const canReadAttachments = isSetSupported("Mailbox", "1.8");
  if (!canReadAttachments && attachments.length > 0) {
    // Ohne Mailbox 1.8 gibt es keinen Weg an die Anhangsinhalte. Die Mail wird trotzdem
    // abgelegt - aber der Benutzer erfaehrt vorher, was fehlen wird.
    warnings.push(T.host.attachmentsUnsupported);
  }

  // Grobschaetzung vor dem Lesen: Anhaenge liegen als Base64 im Speicher, und ein
  // Taskpane-Webview ist kein Ort fuer Ueberraschungen dieser Groessenordnung.
  if (maxBytes > 0 && canReadAttachments) {
    const estimate = attachments.reduce((sum, a) => sum + a.size, 0) * BASE64_OVERHEAD;
    if (estimate > maxBytes) return tooLarge(Math.round(estimate), maxBytes);
  }

  let bodyHtml;
  try {
    bodyHtml = await getBody("html");
  } catch {
    return fail("CLIENT_MIME_FAILED");
  }

  const inlineParts = [];
  const attachedParts = [];
  if (canReadAttachments) {
    for (const attachment of attachments) {
      const part = await readAttachmentPart(attachment, warnings);
      if (!part) continue;
      if (attachment.isInline && attachment.contentId) inlineParts.push(part);
      else attachedParts.push(part);
    }
  }

  const eml = assembleMessage(context, bodyHtml, inlineParts, attachedParts);
  const blob = new Blob([eml], { type: "message/rfc822" });
  if (maxBytes > 0 && blob.size > maxBytes) return tooLarge(blob.size, maxBytes);

  return ok({
    blob,
    bytes: blob.size,
    filename: emlFilename(context.subject),
    subject: context.subject,
    internetMessageId: context.internetMessageId,
    source: "officejs",
    attachmentCount: inlineParts.length + attachedParts.length,
    warnings,
  });
}

/** Anhangsliste ohne Ausnahme - ein fehlendes Element ist kein Grund, alles abzubrechen. */
function safeAttachments() {
  try {
    return getAttachments();
  } catch {
    return [];
  }
}

/** Nur echte Dateianhaenge zaehlen; eingebettete Bilder sind fuer den Benutzer keine. */
function countRealAttachments(attachments) {
  return attachments.filter((a) => !a.isInline).length;
}

/**
 * Liest einen Anhang und macht daraus einen MIME-Teil.
 *
 * Vier Formate liefert Office. Drei davon koennen wir einbetten; `url` ist ein
 * Cloud-Verweis ohne Inhalt - dort gibt es nichts zu uebertragen, und der Benutzer erfaehrt
 * es als Warnung statt als stille Luecke im Ticket.
 */
async function readAttachmentPart(attachment, warnings) {
  let content;
  try {
    content = await getAttachmentContent(attachment.id);
  } catch {
    warnings.push(t("mail.attachmentFailed", { name: attachment.name }));
    return null;
  }

  const format = String(content.format || "").toLowerCase();
  const name = attachment.name || "anhang";

  if (format === "url") {
    warnings.push(t("mail.cloudAttachmentSkipped", { name }));
    return null;
  }

  if (format === "eml") {
    // Eingebettete Nachricht. `message/rfc822` darf laut RFC 2045 nur 7bit/8bit tragen;
    // deshalb hier KEIN Base64 - ein Parser, der daran scheitert, liesse TANSS auf den
    // stillen Dokumenten-Rueckfall zurueckfallen. Preis dieser Wahl: Office liefert die
    // Nachricht bereits als Zeichenkette, wir schreiben sie als UTF-8 - eine eingebettete
    // Nachricht mit abweichender Zeichensatzangabe verliert dabei ihre Treue.
    return {
      headers: [
        "Content-Type: message/rfc822",
        `Content-Disposition: attachment; ${dispositionFilename(name)}`,
        "Content-Transfer-Encoding: 8bit",
      ],
      body: normalizeLineEndings(content.content),
    };
  }

  const contentType =
    format === "icalendar"
      ? "text/calendar; charset=utf-8"
      : attachment.contentType || "application/octet-stream";

  // Base64 kommt bei `base64` bereits fertig aus Office; bei iCalendar ist es Text.
  const base64 = format === "base64" ? content.content : base64OfString(content.content);

  const headers = [
    `Content-Type: ${contentType}; ${nameParameter(name)}`,
    `Content-Disposition: ${attachment.isInline ? "inline" : "attachment"}; ${dispositionFilename(name)}`,
    "Content-Transfer-Encoding: base64",
  ];
  if (attachment.isInline && attachment.contentId) {
    headers.push(`Content-ID: <${attachment.contentId}>`);
  }
  return { headers, body: wrapBase64(base64) };
}

/**
 * Setzt die Nachricht zusammen.
 *
 * Aufbau je nach Inhalt, damit `cid:`-Verweise im Text aufloesbar bleiben:
 *   nur Text                      -> text/html
 *   Text + eingebettete Bilder    -> multipart/related
 *   dazu echte Anhaenge           -> multipart/mixed { related, anhaenge… }
 */
function assembleMessage(context, bodyHtml, inlineParts, attachedParts) {
  const bodyPart = {
    headers: ['Content-Type: text/html; charset="utf-8"', "Content-Transfer-Encoding: base64"],
    body: wrapBase64(base64OfString(bodyHtml || "")),
  };

  let root = bodyPart;
  if (inlineParts.length > 0) {
    root = multipart("related", [bodyPart, ...inlineParts], 'type="text/html"');
  }
  if (attachedParts.length > 0) {
    root = multipart("mixed", [root, ...attachedParts]);
  }

  const headers = [
    "MIME-Version: 1.0",
    `Date: ${formatDate(context.dateTimeCreated)}`,
    `From: ${formatAddress(context.from)}`,
  ];
  if (context.to.length > 0) headers.push(`To: ${formatAddressList(context.to)}`);
  if (context.cc.length > 0) headers.push(`Cc: ${formatAddressList(context.cc)}`);
  headers.push(`Subject: ${encodeHeaderText(context.subject || "")}`);
  if (context.internetMessageId) headers.push(`Message-ID: ${context.internetMessageId}`);

  return [...headers, ...root.headers, "", root.body].join(CRLF) + CRLF;
}

/** Baut einen Mehrteiler und erzeugt dafuer eine Grenzmarke, die im Inhalt nicht vorkommt. */
function multipart(subtype, parts, extraParameter = "") {
  const boundary = `----tanss-${randomToken()}`;
  const lines = [];
  for (const part of parts) {
    lines.push(`--${boundary}`, ...part.headers, "", part.body);
  }
  lines.push(`--${boundary}--`);
  return {
    headers: [
      `Content-Type: multipart/${subtype}; boundary="${boundary}"${extraParameter ? `; ${extraParameter}` : ""}`,
    ],
    body: lines.join(CRLF),
  };
}

/* --------------------------------------------------------------- Kodierungen */

/** Zufallsmarke fuer Grenzen. Aus dem Systemzufall, damit sie nie erratbar kollidiert. */
function randomToken() {
  const bytes = new Uint8Array(12);
  globalThis.crypto.getRandomValues(bytes);
  return Array.from(bytes, (b) => b.toString(16).padStart(2, "0")).join("");
}

function base64OfBytes(bytes) {
  let binary = "";
  const chunkSize = 0x8000;
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + chunkSize));
  }
  return btoa(binary);
}

function base64OfString(value) {
  return base64OfBytes(new TextEncoder().encode(String(value)));
}

/** Base64 auf 76 Zeichen umbrechen - laengere Zeilen verletzen RFC 2045. */
function wrapBase64(base64) {
  const clean = String(base64).replace(/[\r\n]/g, "");
  const lines = [];
  for (let offset = 0; offset < clean.length; offset += 76) {
    lines.push(clean.slice(offset, offset + 76));
  }
  return lines.join(CRLF);
}

/** Alle Zeilenenden auf CRLF bringen. Gemischte Enden bringen manche Parser aus dem Tritt. */
function normalizeLineEndings(value) {
  return String(value).replace(/\r\n|\r|\n/g, CRLF);
}

/** Ob der Text ohne Kodierung in eine Kopfzeile darf - druckbares US-ASCII, sonst nichts. */
function isAscii(value) {
  return /^[\x20-\x7e]*$/.test(value);
}

/**
 * Kopfzeilentext nach RFC 2047, wenn er nicht rein ASCII ist.
 *
 * Ein Betreff mit Umlaut roh in die Kopfzeile zu schreiben, ergibt in TANSS entweder
 * Zeichensalat oder eine abgewiesene Nachricht. Die Stuecke bleiben unter 75 Zeichen je
 * kodiertem Wort, wie es die Norm verlangt.
 */
function encodeHeaderText(value) {
  const text = String(value).replace(/[\r\n]+/g, " ").trim();
  if (text === "") return "";
  if (isAscii(text)) return text;

  const encoder = new TextEncoder();
  const words = [];
  let chunk = "";
  let used = 0;
  // Geschnitten wird an ZEICHENgrenzen, nie an Bytegrenzen: ein ueber zwei kodierte
  // Woerter zerrissenes Mehrbytezeichen - RFC 2047 Abschnitt 5 verbietet es ausdruecklich
  // - kommt in TANSS als Ersetzungszeichen an, aus "über Nacht" wird "?ber Nacht".
  // `for...of` laeuft ueber Codepunkte und haelt damit auch Ersatzzeichenpaare zusammen.
  for (const character of text) {
    const size = encoder.encode(character).length;
    // 45 Rohbytes ergeben 60 Base64-Zeichen; mit `=?utf-8?B?…?=` bleibt das unter 75.
    if (used + size > 45 && chunk !== "") {
      words.push(encodedWord(chunk));
      chunk = "";
      used = 0;
    }
    chunk += character;
    used += size;
  }
  if (chunk !== "") words.push(encodedWord(chunk));
  return words.join(`${CRLF} `);
}

/** Ein einzelnes kodiertes Wort nach RFC 2047. */
function encodedWord(text) {
  return `=?utf-8?B?${base64OfString(text)}?=`;
}

/** Adressangabe `"Name" <adresse>`; der Name wird kodiert, die Adresse nie veraendert. */
function formatAddress(address) {
  if (!address || !address.address) return "unknown@invalid";
  const name = String(address.name || "").trim();
  if (name === "") return `<${address.address}>`;
  if (isAscii(name)) {
    return `"${name.replace(/["\\]/g, "\\$&")}" <${address.address}>`;
  }
  // Kodierte Woerter duerfen NICHT in Anfuehrungszeichen stehen - sonst werden sie
  // woertlich genommen und der Empfaenger sieht `=?utf-8?B?…?=`.
  return `${encodeHeaderText(name)} <${address.address}>`;
}

function formatAddressList(addresses) {
  return addresses.map(formatAddress).join(`,${CRLF} `);
}

/**
 * Datum nach RFC 5322.
 *
 * Von Hand gebaut statt ueber `toUTCString()`: dort steht am Ende `GMT`, was nur noch die
 * veraltete Schreibweise ist. `+0000` versteht jeder Parser.
 */
function formatDate(value) {
  const date = value instanceof Date && !Number.isNaN(value.getTime()) ? value : new Date();
  const days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  const months = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
  ];
  const pad = (n) => String(n).padStart(2, "0");
  return (
    `${days[date.getUTCDay()]}, ${pad(date.getUTCDate())} ${months[date.getUTCMonth()]} ` +
    `${date.getUTCFullYear()} ${pad(date.getUTCHours())}:${pad(date.getUTCMinutes())}:` +
    `${pad(date.getUTCSeconds())} +0000`
  );
}

/**
 * Dateinamen-Parameter.
 *
 * Entweder `filename` in ASCII, wenn der Name ohnehin ASCII ist, oder `filename*` nach
 * RFC 2231 - nie beides nebeneinander. Stehen beide in derselben Kopfzeile, nimmt ein
 * Parser den ASCII-Ersatz, und der Anhang liegt in TANSS unter einem Namen voller
 * Unterstriche.
 */
function dispositionFilename(name) {
  return parameterPair("filename", name);
}

function nameParameter(name) {
  return parameterPair("name", name);
}

function parameterPair(key, name) {
  const clean = String(name).replace(/[\r\n"\\]/g, "_").slice(0, 120);
  const ascii = clean.replace(/[^\x20-\x7e]/g, "_");
  if (clean === ascii) return `${key}="${ascii}"`;
  // Nur die kodierte Form: steht der ASCII-Ersatz daneben, liefert Pythons `email` -
  // dieselbe Bibliothek, mit der der Dienst die hochgeladene .eml vorprueft - ihn statt
  // des echten Namens, und zwar unabhaengig von der Reihenfolge der Parameter. Der
  // Anhang hiesse in TANSS dann "M_ller & S_hne.pdf".
  return `${key}*=UTF-8''${encodeRfc2231(clean)}`;
}

function encodeRfc2231(value) {
  return Array.from(new TextEncoder().encode(value), (byte) =>
    (byte >= 0x30 && byte <= 0x39) ||
    (byte >= 0x41 && byte <= 0x5a) ||
    (byte >= 0x61 && byte <= 0x7a) ||
    byte === 0x2d ||
    byte === 0x2e ||
    byte === 0x5f
      ? String.fromCharCode(byte)
      : `%${byte.toString(16).toUpperCase().padStart(2, "0")}`,
  ).join("");
}

/**
 * Dateiname der abgelegten `.eml`.
 *
 * Nur ein Vorschlag - der Dienst bereinigt noch einmal und ist die massgebliche Stelle.
 * Hier wird trotzdem bereinigt, damit kein Pfadtrenner und kein Steuerzeichen ueberhaupt
 * erst auf die Reise geht.
 */
export function emlFilename(subject) {
  const clean = String(subject || "")
    .replace(/[\u0000-\u001f\u007f]/g, " ")
    .replace(/[\\/:*?"<>|]/g, "-")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 80);
  return `${clean === "" ? "mail" : clean}.eml`;
}

function tooLarge(actual, limit) {
  return {
    ok: false,
    error: {
      code: "CLIENT_TOO_LARGE",
      message: t("mail.tooLargeLocal", {
        size: formatBytes(actual),
        limit: formatBytes(limit),
      }),
    },
  };
}
